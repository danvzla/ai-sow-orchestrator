#!/usr/bin/env python3
"""
╔═══════════════════════════════════════════════════════════════╗
║          SOW Orchestrator — Python CLI                        ║
║          AgentPM AI Portfolio · Daniel Mazzini                ║
║                                                               ║
║  Autonomous 3-agent pipeline:                                 ║
║    Agent 1 · Discovery Analyst  — structures raw notes        ║
║    Agent 2 · Scope Validator    — autonomous go/no-go gate    ║
║    Agent 3 · SOW Drafter        — writes SOW if approved      ║
╚═══════════════════════════════════════════════════════════════╝
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

# ─── DEPENDENCY CHECKS ────────────────────────────────────────────────────────

try:
    import anthropic
except ImportError:
    print("\n[ERROR] anthropic package not found.")
    print("Run: pip install anthropic\n")
    sys.exit(1)

try:
    from rich.console import Console
    from rich.panel import Panel
    from rich.progress import Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
    from rich.table import Table
    from rich.text import Text
    from rich.rule import Rule
    from rich.columns import Columns
    from rich.align import Align
    HAS_RICH = True
    console = Console()
except ImportError:
    HAS_RICH = False
    print("Tip: Install 'rich' for better output: pip install rich\n")
    class Console:
        def print(self, *args, **kwargs): print(*args)
        def log(self, *args, **kwargs): print(*args)
    console = Console()

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ─── CONSTANTS ────────────────────────────────────────────────────────────────

VERSION   = "1.0.0"
MODEL     = "claude-haiku-4-5-20251001"
CHARCOAL  = "#2C2C2C"
GOLD      = "#C9A84C"

ENGAGEMENT_TYPES = [
    "Private Cloud Deployment (VCF)",
    "Network Security Hardening (NSX)",
    "Telco NFV Transformation",
    "SASE / SD-WAN Migration",
    "Cloud Modernization Assessment",
    "NOC Automation Program",
    "Data Center Consolidation",
    "Hybrid Cloud Architecture",
]

ENGAGEMENT_SIZES = [
    "Small (under $150K)",
    "Medium ($150K–$500K)",
    "Large ($500K–$2M)",
    "Enterprise ($2M+)",
]

INDUSTRIES = [
    "Financial Services", "Healthcare", "Telco / CSP",
    "Manufacturing", "Technology", "Government",
    "Energy", "Retail", "Not specified",
]

# ─── JSON REPAIR ──────────────────────────────────────────────────────────────

def repair_json(raw: str) -> dict:
    """4-attempt JSON recovery chain."""
    first = raw.find('{')
    last  = raw.rfind('}')
    if first == -1 or last == -1:
        raise ValueError("No JSON object found in response")
    chunk = raw[first:last+1]

    for attempt, transform in enumerate([
        lambda s: s,
        lambda s: re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', s),
        lambda s: re.sub(r'(?<!\\)\n', '\\n', s),
        lambda s: re.sub(r'\r?\n|\t', ' ', s),
    ], start=1):
        try:
            return json.loads(transform(chunk))
        except json.JSONDecodeError:
            continue

    raise ValueError("All JSON repair attempts failed — please try again.")

# ─── SANITISE NOTES ───────────────────────────────────────────────────────────

def sanitise(text: str) -> str:
    """Remove characters that break JSON string values."""
    return (
        text.replace('\\', '/')
            .replace('"', "'")
            .replace('\r\n', ' | ')
            .replace('\n', ' | ')
            .replace('\t', ' ')
    )

# ─── AGENT 1: DISCOVERY ANALYST ───────────────────────────────────────────────

def run_agent1(client, notes: str, eng_type: str, eng_size: str, industry: str) -> dict:
    safe_notes = sanitise(notes)
    prompt = f"""You are a Senior Solutions Architect and TPM with 20+ years structuring professional services engagements.

Read these discovery notes and extract a structured requirements brief.
Return ONLY raw JSON — no markdown, no backticks. Start with {{ end with }}.
All string values must be on ONE LINE — no line breaks inside strings.

ENGAGEMENT TYPE: {eng_type}
ENGAGEMENT SIZE: {eng_size}
CUSTOMER INDUSTRY: {industry}

DISCOVERY NOTES:
{safe_notes}

Return this JSON — keep all string values SHORT and on ONE LINE:
{{
  "engagement_objective": "one sentence describing success",
  "client_profile": "two sentences describing customer and context",
  "in_scope": ["workstream 1", "workstream 2", "workstream 3"],
  "out_of_scope": ["excluded item 1", "excluded item 2"],
  "deliverables": [
    {{"name": "deliverable name", "description": "what it includes", "format": "document or workshop or implementation"}}
  ],
  "phases": [
    {{"phase": "phase name", "duration": "e.g. 4 weeks", "activities": ["activity 1", "activity 2"]}}
  ],
  "assumptions": ["assumption 1", "assumption 2"],
  "customer_responsibilities": ["responsibility 1", "responsibility 2"],
  "dependencies": ["dependency 1"],
  "success_criteria": ["criterion 1", "criterion 2"],
  "risks": [
    {{"risk": "risk title", "mitigation": "how to address it"}}
  ],
  "estimated_duration": "total duration e.g. 14 weeks",
  "team_structure": "recommended team composition",
  "open_items": ["open item 1"]
}}"""

    message = client.messages.create(
        model=MODEL,
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = message.content[0].text
    return repair_json(raw)

# ─── AGENT 2: SCOPE VALIDATOR ─────────────────────────────────────────────────

def run_agent2(client, brief: dict, eng_type: str, eng_size: str, industry: str) -> dict:
    prompt = f"""You are a Senior QA Architect reviewing a requirements brief before SOW drafting.

Check for: contradictions, missing critical items, ambiguities, unrealistic timelines, scope gaps.
Return ONLY raw JSON — no markdown. Start with {{ end with }}.
All string values on ONE LINE. Max 4 issues, 4 warnings, 4 approved_items.

ENGAGEMENT TYPE: {eng_type}
ENGAGEMENT SIZE: {eng_size}
CUSTOMER INDUSTRY: {industry}

BRIEF:
{json.dumps(brief)}

Return:
{{
  "decision": "approved",
  "validation_summary": "one sentence summary",
  "quality_score": 85,
  "issues": [
    {{"type": "Contradiction", "title": "short title", "description": "problem", "recommendation": "fix"}}
  ],
  "warnings": [
    {{"type": "Risk", "title": "short title", "description": "watch for this"}}
  ],
  "approved_items": ["solid item 1", "solid item 2"]
}}

Decision: "rejected" if critical issues, "approved_with_warnings" if only warnings, "approved" if clean."""

    message = client.messages.create(
        model=MODEL,
        max_tokens=6000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = message.content[0].text
    return repair_json(raw)

# ─── AGENT 3: SOW DRAFTER ─────────────────────────────────────────────────────

def run_agent3(client, brief: dict, validation: dict, eng_type: str, eng_size: str, industry: str) -> str:
    warnings_ctx = ""
    if validation.get("warnings"):
        warnings_ctx = "\n\nVALIDATION WARNINGS (address in SOW):\n"
        warnings_ctx += "\n".join(f"- {w['title']}: {w['description']}" for w in validation["warnings"])

    prompt = f"""You are a Senior Principal Consultant with 20+ years writing professional SOW documents for VMware, Dell EMC, Juniper Networks, and Nokia/Ericsson. This brief has been validated and approved.

Write a Statement of Work in clean Markdown. Use tables wherever possible. Each section opens with 1-2 sentences of professional prose, then a table. Be specific — reference actual technologies and customer context.

ENGAGEMENT TYPE: {eng_type}
ENGAGEMENT SIZE: {eng_size}
CUSTOMER INDUSTRY: {industry}

VALIDATED BRIEF:
{json.dumps(brief)}{warnings_ctx}

Write exactly these sections in Markdown:

# [Engagement Type] — Statement of Work
**Version:** 1.0 | **Date:** {datetime.today().strftime('%B %d, %Y')} | **Status:** Draft for Review

## 1. Engagement Overview
[2 sentences on customer context]
| Field | Detail |
[table with: Engagement Type, Industry, Size, Duration, Delivery Model, Engagement Lead]

## 2. Engagement Objective
[2 specific sentences referencing actual technologies and business outcome]

## 3. Scope of Work
[1 sentence]
| Workstream | Description | Key Activities | In Scope |
[one row per workstream]

## 4. Deliverables
[1 sentence]
| # | Deliverable | Description | Format | Qty | Est. Effort (hrs) |
[each deliverable with realistic hours]
| **TOTAL** | | | | | **XX hrs** |

## 5. Work Breakdown & Schedule
[1 sentence]
| Phase | Activity | Owner | Duration | Hours |
[grouped by phase with subtotals. Final row: TOTAL XX hrs]

## 6. Milestones & Acceptance
[1 sentence]
| # | Milestone | Acceptance Criteria | Target Week | Owner |

## 7. Assumptions
[1 sentence]
| # | Assumption | Dependency | Risk if Invalid |

## 8. Customer Responsibilities
[1 sentence]
| # | Responsibility | Details | Required By | Owner |

## 9. Out of Scope
[1 sentence]
| # | Item | Rationale | How to Add |

## 10. Risk Register
[1 sentence]
| # | Risk | Probability | Impact | Mitigation | Owner |

## 11. Team Structure
[1 sentence]
| Role | Responsibilities | Allocation | Location |

## 12. Open Items
[1 sentence]
| # | Item | Decision Required | Owner | Due Date |

## 13. Commercial Terms
[1 sentence]
| Term | Details |
[rows: Engagement Model, Payment Schedule, Travel & Expenses, Change Order Process, Cancellation Notice, Warranty Period]

## 14. Document Control
| Version | Date | Author | Changes |
| 1.0 | {datetime.today().strftime('%Y-%m-%d')} | Daniel Mazzini | Initial draft |

---
*This SOW is a draft generated by the AgentPM SOW Orchestrator. Subject to legal review and customer negotiation before execution.*"""

    message = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text

# ─── RICH DISPLAY HELPERS ─────────────────────────────────────────────────────

def print_header():
    if not HAS_RICH:
        print("\n=== SOW Orchestrator — AgentPM ===\n")
        return
    console.print()
    console.print(Panel.fit(
        "[bold #C9A84C]SOW Orchestrator[/bold #C9A84C]  ·  [dim]AgentPM AI Portfolio · Daniel Mazzini[/dim]\n"
        "[dim]3-Agent Autonomous Pipeline · Discovery Analysis · Scope Validation · SOW Drafting[/dim]",
        border_style="#C9A84C",
        padding=(1, 4),
    ))
    console.print()

def print_agent_start(num: int, name: str, desc: str):
    if not HAS_RICH:
        print(f"\n[Agent {num}] {name} — {desc}")
        return
    colors = {1: "#C9A84C", 2: "#854F0B", 3: "#3B6D11"}
    color = colors.get(num, "white")
    console.print(Rule(f"[bold {color}]Agent {num} · {name}[/bold {color}]", style=color))
    console.print(f"  [dim]{desc}[/dim]")

def print_agent_done(num: int, summary: str):
    if not HAS_RICH:
        print(f"  ✓ {summary}")
        return
    console.print(f"  [green]✓[/green] {summary}")

def print_agent_error(num: int, msg: str):
    if not HAS_RICH:
        print(f"  ✗ Error: {msg}")
        return
    console.print(f"  [red]✗ Error:[/red] {msg}")

def print_brief_summary(brief: dict):
    if not HAS_RICH:
        print(f"\nBrief: {len(brief.get('in_scope',[]))} workstreams · "
              f"{len(brief.get('deliverables',[]))} deliverables · "
              f"{len(brief.get('risks',[]))} risks")
        return
    t = Table(show_header=False, box=None, padding=(0,2))
    t.add_column(style="dim", width=22)
    t.add_column()
    t.add_row("Objective",    brief.get("engagement_objective","")[:80])
    t.add_row("Workstreams",  str(len(brief.get("in_scope",[]))))
    t.add_row("Deliverables", str(len(brief.get("deliverables",[]))))
    t.add_row("Phases",       str(len(brief.get("phases",[]))))
    t.add_row("Risks",        str(len(brief.get("risks",[]))))
    t.add_row("Duration",     brief.get("estimated_duration","—"))
    console.print(Panel(t, title="[bold]Agent 1 Output — Requirements Brief[/bold]", border_style="dim", padding=(0,1)))

def print_validation_result(v: dict):
    decision = v.get("decision","approved")
    score    = v.get("quality_score", 0)
    issues   = v.get("issues", [])
    warnings = v.get("warnings", [])

    color_map = {"approved":"green", "approved_with_warnings":"yellow", "rejected":"red"}
    icon_map  = {"approved":"✓", "approved_with_warnings":"⚠", "rejected":"✗"}
    color     = color_map.get(decision, "white")
    icon      = icon_map.get(decision, "?")

    if HAS_RICH:
        label = decision.replace("_", " ").upper()
        console.print(Panel(
            f"[bold {color}]{icon} {label}[/bold {color}]  ·  Quality score: {score}/100\n"
            f"[dim]{v.get('validation_summary','')}[/dim]",
            title="[bold]Agent 2 Output — Scope Validation[/bold]",
            border_style=color,
            padding=(0,1),
        ))
        if issues:
            t = Table("Type","Title","Fix", box=None, padding=(0,2))
            for i in issues:
                t.add_row(f"[red]{i.get('type','')}[/red]", i.get('title',''), i.get('recommendation',''))
            console.print(t)
        if warnings:
            t = Table("Type","Title","Description", box=None, padding=(0,2))
            for w in warnings:
                t.add_row(f"[yellow]{w.get('type','')}[/yellow]", w.get('title',''), w.get('description',''))
            console.print(t)
    else:
        print(f"\nValidation: {decision.upper()} (score: {score}/100)")
        print(v.get("validation_summary",""))
        for i in issues:
            print(f"  ISSUE [{i.get('type','')}]: {i.get('title','')} — Fix: {i.get('recommendation','')}")
        for w in warnings:
            print(f"  WARNING [{w.get('type','')}]: {w.get('title','')} — {w.get('description','')}")

# ─── DOCX EXPORT ──────────────────────────────────────────────────────────────

def hex_to_rgb(hex_color: str):
    h = hex_color.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def set_cell_bg(cell, hex_color: str):
    """Set table cell background color."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color.lstrip('#'))
    tcPr.append(shd)

def export_docx(sow_markdown: str, output_path: Path, eng_type: str):
    """Convert SOW markdown to a formatted Word document."""
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    charcoal_rgb = hex_to_rgb("2C2C2C")
    gold_rgb     = hex_to_rgb("C9A84C")
    white_rgb    = (255, 255, 255)

    lines = sow_markdown.split('\n')

    def add_heading(text: str, level: int):
        p = doc.add_heading(text.strip(), level=min(level, 3))
        run = p.runs[0] if p.runs else p.add_run(text.strip())
        if level == 1:
            run.font.color.rgb = RGBColor(*charcoal_rgb)
            run.font.size = Pt(20)
            run.font.bold = True
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(6)
            # Gold underline border
            pPr    = p._p.get_or_add_pPr()
            pBdr   = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'),   'single')
            bottom.set(qn('w:sz'),    '12')
            bottom.set(qn('w:space'), '4')
            bottom.set(qn('w:color'), 'C9A84C')
            pBdr.append(bottom)
            pPr.append(pBdr)
        elif level == 2:
            run.font.color.rgb = RGBColor(*white_rgb)
            run.font.size = Pt(11)
            run.font.bold = True
            # Charcoal background shading
            pPr  = p._p.get_or_add_pPr()
            shd  = OxmlElement('w:shd')
            shd.set(qn('w:val'),   'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'),  '2C2C2C')
            pPr.append(shd)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after  = Pt(4)

    def add_paragraph(text: str):
        if not text.strip():
            return
        # Handle bold **text**
        p   = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(4)
        parts = re.split(r'\*\*(.*?)\*\*', text.strip())
        for i, part in enumerate(parts):
            run = p.add_run(part)
            run.font.name  = "Calibri"
            run.font.size  = Pt(11)
            run.font.color.rgb = RGBColor(*charcoal_rgb)
            if i % 2 == 1:
                run.font.bold = True

    def add_table_from_markdown(md_rows: list):
        """Build a formatted DOCX table from markdown table rows."""
        rows = [r for r in md_rows if r.strip().startswith('|') and not re.match(r'^\|[-| :]+\|$', r.strip())]
        if len(rows) < 2:
            return

        data = []
        for row in rows:
            cells = [c.strip() for c in row.strip().strip('|').split('|')]
            data.append(cells)

        if not data:
            return

        max_cols = max(len(r) for r in data)
        t = doc.add_table(rows=len(data), cols=max_cols)
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.LEFT

        for r_idx, row in enumerate(data):
            for c_idx, cell_text in enumerate(row[:max_cols]):
                cell = t.cell(r_idx, c_idx)
                clean = re.sub(r'\*\*(.*?)\*\*', r'\1', cell_text.strip())
                p = cell.paragraphs[0]
                run = p.add_run(clean)
                run.font.name = "Calibri"
                run.font.size = Pt(10)
                if r_idx == 0:
                    # Header row: charcoal bg, white text, bold
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(*white_rgb)
                    set_cell_bg(cell, "2C2C2C")
                elif clean.lower().startswith("total") or clean.lower() == "**total**":
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(*charcoal_rgb)
                    set_cell_bg(cell, "F5E9C8")
                elif r_idx % 2 == 0:
                    run.font.color.rgb = RGBColor(*charcoal_rgb)
                    set_cell_bg(cell, "F8FAFC")
                else:
                    run.font.color.rgb = RGBColor(*charcoal_rgb)

        doc.add_paragraph()

    # Parse and render markdown
    i = 0
    table_buffer = []

    while i < len(lines):
        line = lines[i]

        # H1
        if line.startswith('# ') and not line.startswith('## '):
            if table_buffer:
                add_table_from_markdown(table_buffer)
                table_buffer = []
            add_heading(line[2:], 1)

        # H2
        elif line.startswith('## '):
            if table_buffer:
                add_table_from_markdown(table_buffer)
                table_buffer = []
            add_heading(line[3:], 2)

        # H3
        elif line.startswith('### '):
            if table_buffer:
                add_table_from_markdown(table_buffer)
                table_buffer = []
            add_heading(line[4:], 3)

        # Table row
        elif line.strip().startswith('|'):
            table_buffer.append(line)

        # Separator line (---)
        elif re.match(r'^[-─]+$', line.strip()):
            if table_buffer:
                add_table_from_markdown(table_buffer)
                table_buffer = []

        # Regular paragraph
        elif line.strip():
            if table_buffer:
                add_table_from_markdown(table_buffer)
                table_buffer = []
            if not line.startswith('*This SOW'):
                add_paragraph(line)

        # Blank line — flush table if pending
        else:
            if table_buffer:
                add_table_from_markdown(table_buffer)
                table_buffer = []

        i += 1

    if table_buffer:
        add_table_from_markdown(table_buffer)

    doc.save(str(output_path))

# ─── INTERACTIVE INPUT ────────────────────────────────────────────────────────

def select_from_list(prompt: str, options: list, default: int = 0) -> str:
    if HAS_RICH:
        console.print(f"\n[bold]{prompt}[/bold]")
        for i, opt in enumerate(options, 1):
            marker = "[gold1]→[/gold1]" if i-1 == default else " "
            console.print(f"  {marker} [{i}] {opt}")
        choice = console.input(f"  [dim]Enter number (default {default+1}): [/dim]").strip()
    else:
        print(f"\n{prompt}")
        for i, opt in enumerate(options, 1):
            print(f"  [{i}] {opt}")
        choice = input(f"  Enter number (default {default+1}): ").strip()

    if not choice:
        return options[default]
    try:
        idx = int(choice) - 1
        return options[max(0, min(idx, len(options)-1))]
    except ValueError:
        return options[default]

# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="SOW Orchestrator — Autonomous 3-agent SOW generation pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python sow_generator.py notes.txt
  python sow_generator.py notes.txt --type "Telco NFV Transformation" --industry "Telco / CSP"
  python sow_generator.py notes.txt --no-docx --output my_sow.md
  python sow_generator.py --interactive

AgentPM AI Portfolio · Daniel Mazzini · github.com/danvzla
        """
    )
    parser.add_argument("notes_file",     nargs="?",  help="Path to discovery notes text file")
    parser.add_argument("--type",         default=None, help="Engagement type")
    parser.add_argument("--size",         default=None, help="Engagement size")
    parser.add_argument("--industry",     default=None, help="Customer industry")
    parser.add_argument("--output",       default=None, help="Output file path (without extension)")
    parser.add_argument("--no-docx",      action="store_true", help="Skip Word document export")
    parser.add_argument("--interactive",  action="store_true", help="Interactive mode — prompts for all inputs")
    parser.add_argument("--api-key",      default=None, help="Anthropic API key (or set ANTHROPIC_API_KEY env var)")
    args = parser.parse_args()

    print_header()

    # API key
    api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY","")
    if not api_key:
        if HAS_RICH:
            api_key = console.input("[bold]Anthropic API key:[/bold] ").strip()
        else:
            api_key = input("Anthropic API key: ").strip()
    if not api_key:
        print("Error: No API key provided. Set ANTHROPIC_API_KEY or use --api-key.")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    # Discovery notes
    if args.notes_file:
        path = Path(args.notes_file)
        if not path.exists():
            print(f"Error: File not found: {path}")
            sys.exit(1)
        notes = path.read_text(encoding="utf-8").strip()
        if HAS_RICH:
            console.print(f"[dim]Notes loaded from:[/dim] {path} ({len(notes)} chars)")
        else:
            print(f"Notes loaded: {path} ({len(notes)} chars)")
    elif args.interactive or not args.notes_file:
        if HAS_RICH:
            console.print("\n[bold]Paste your discovery notes below.[/bold]")
            console.print("[dim]Type END on a new line when done:[/dim]\n")
        else:
            print("\nPaste discovery notes. Type END when done:\n")
        lines = []
        while True:
            try:
                line = input()
                if line.strip().upper() == "END":
                    break
                lines.append(line)
            except EOFError:
                break
        notes = "\n".join(lines).strip()
        if not notes:
            print("Error: No discovery notes provided.")
            sys.exit(1)
    else:
        print("Error: Provide a notes file or use --interactive.")
        parser.print_help()
        sys.exit(1)

    # Engagement context
    if args.interactive or not args.type:
        eng_type = select_from_list("Engagement type:", ENGAGEMENT_TYPES)
    else:
        eng_type = args.type

    if args.interactive or not args.size:
        eng_size = select_from_list("Engagement size:", ENGAGEMENT_SIZES, default=1)
    else:
        eng_size = args.size

    if args.interactive or not args.industry:
        industry = select_from_list("Customer industry:", INDUSTRIES)
    else:
        industry = args.industry

    if HAS_RICH:
        console.print(f"\n[dim]Engagement:[/dim] {eng_type}  ·  {eng_size}  ·  {industry}\n")

    # Output path
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_stem = args.output or f"SOW_{eng_type.replace(' ','_').replace('/','_').replace('(','').replace(')','')[:30]}_{timestamp}"
    output_md   = Path(output_stem + ".md")
    output_docx = Path(output_stem + ".docx")

    # ── AGENT 1 ───────────────────────────────────────────────────────────────
    print_agent_start(1, "Discovery Analyst", "Reading notes · structuring scope · mapping deliverables and risks")

    try:
        if HAS_RICH:
            with Progress(SpinnerColumn(), TextColumn("[dim]{task.description}[/dim]"), TimeElapsedColumn(), console=console) as prog:
                task = prog.add_task("Analyzing discovery notes...", total=None)
                brief = run_agent1(client, notes, eng_type, eng_size, industry)
        else:
            print("  Running Agent 1...")
            brief = run_agent1(client, notes, eng_type, eng_size, industry)
    except Exception as e:
        print_agent_error(1, str(e))
        sys.exit(1)

    print_agent_done(1, f"Brief complete · {len(brief.get('in_scope',[]))} workstreams · {len(brief.get('deliverables',[]))} deliverables · {len(brief.get('risks',[]))} risks · duration: {brief.get('estimated_duration','—')}")
    print_brief_summary(brief)

    # ── AGENT 2 ───────────────────────────────────────────────────────────────
    print_agent_start(2, "Scope Validator", "Checking for contradictions · gaps · ambiguities · autonomous decision")

    try:
        if HAS_RICH:
            with Progress(SpinnerColumn(), TextColumn("[dim]{task.description}[/dim]"), TimeElapsedColumn(), console=console) as prog:
                task = prog.add_task("Validating scope brief...", total=None)
                validation = run_agent2(client, brief, eng_type, eng_size, industry)
        else:
            print("  Running Agent 2...")
            validation = run_agent2(client, brief, eng_type, eng_size, industry)
    except Exception as e:
        print_agent_error(2, str(e))
        sys.exit(1)

    decision = validation.get("decision","approved")
    print_validation_result(validation)

    if decision == "rejected":
        print_agent_error(2, "Scope rejected. Fix the issues above and rerun.")
        if HAS_RICH:
            console.print("\n[red]Agent 3 blocked.[/red] Edit your discovery notes and run again.\n")
        else:
            print("\nAgent 3 blocked. Edit discovery notes and rerun.")
        sys.exit(2)

    print_agent_done(2, f"Decision: {decision.upper()} · Quality score: {validation.get('quality_score',0)}/100")

    # ── AGENT 3 ───────────────────────────────────────────────────────────────
    print_agent_start(3, "SOW Drafter", "Writing 14-section Statement of Work with tables and effort estimates")

    try:
        if HAS_RICH:
            with Progress(SpinnerColumn(), TextColumn("[dim]{task.description}[/dim]"), TimeElapsedColumn(), console=console) as prog:
                task = prog.add_task("Drafting Statement of Work...", total=None)
                sow_markdown = run_agent3(client, brief, validation, eng_type, eng_size, industry)
        else:
            print("  Running Agent 3...")
            sow_markdown = run_agent3(client, brief, validation, eng_type, eng_size, industry)
    except Exception as e:
        print_agent_error(3, str(e))
        sys.exit(1)

    # Save markdown
    output_md.write_text(sow_markdown, encoding="utf-8")
    print_agent_done(3, f"SOW drafted · {len(sow_markdown.split(chr(10)))} lines · saved to {output_md}")

    # Save DOCX
    if not args.no_docx:
        if HAS_DOCX:
            try:
                export_docx(sow_markdown, output_docx, eng_type)
                if HAS_RICH:
                    console.print(f"  [green]✓[/green] Word document saved to [bold]{output_docx}[/bold]")
                else:
                    print(f"  ✓ Word document saved to {output_docx}")
            except Exception as e:
                if HAS_RICH:
                    console.print(f"  [yellow]⚠ DOCX export failed:[/yellow] {e} (Markdown saved successfully)")
                else:
                    print(f"  ⚠ DOCX export failed: {e}")
        else:
            if HAS_RICH:
                console.print("  [yellow]⚠ python-docx not installed — skipping Word export.[/yellow]")
                console.print("    Run: [bold]pip install python-docx[/bold]")
            else:
                print("  ⚠ python-docx not installed. Run: pip install python-docx")

    # ── SUMMARY ───────────────────────────────────────────────────────────────
    if HAS_RICH:
        console.print()
        console.print(Panel(
            f"[bold green]✓ SOW generation complete[/bold green]\n\n"
            f"[dim]Markdown:[/dim]  {output_md}\n"
            f"[dim]Word doc:[/dim]  {output_docx if (not args.no_docx and HAS_DOCX) else 'skipped'}\n\n"
            f"[dim]Quality score:[/dim] {validation.get('quality_score',0)}/100  ·  "
            f"[dim]Validation:[/dim] {decision.upper()}  ·  "
            f"[dim]Duration:[/dim] {brief.get('estimated_duration','—')}",
            border_style="green",
            padding=(1, 2),
        ))
        console.print()
        console.print("[dim]AgentPM SOW Orchestrator · Daniel Mazzini · github.com/danvzla[/dim]\n")
    else:
        print(f"\n✓ Complete. Markdown: {output_md}")
        if not args.no_docx and HAS_DOCX:
            print(f"✓ Word doc: {output_docx}")


if __name__ == "__main__":
    main()
