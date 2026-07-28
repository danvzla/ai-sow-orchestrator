# AI SOW Orchestrator

> **3-Agent AI System** — Turns raw discovery notes into a structured, validation-gated Statement of Work draft.

**Live demo →** [danvzla.github.io/ai-sow-orchestrator](https://danvzla.github.io/ai-sow-orchestrator)

Part of the **AI Portfolio** by [Daniel Mazzini](https://www.linkedin.com/in/daniel-mazzini-22059734/) — Principal Architect & Senior TPM

---

## What it does

Paste your raw discovery notes — bullet points, meeting notes, stream of consciousness — and a 3-agent pipeline autonomously produces a complete SOW:

| Agent | Role | Output |
|---|---|---|
| **Agent 1 · Discovery Analyst** | Reads raw notes, structures scope, maps deliverables, identifies risks | Structured requirements brief (JSON) |
| **Agent 2 · Scope Validator** | Reviews brief for contradictions, gaps, ambiguities — autonomous go/no-go decision | Validation report with quality score |
| **Agent 3 · SOW Drafter** | Runs **only if Agent 2 approves** · Writes complete 14-section SOW | Full Statement of Work |

Agent 3 is **gated** — if Agent 2 rejects the scope, Agent 3 is blocked and no SOW is produced. This enforces quality before any drafting begins.

---

## Three Modes

| Mode | Description | Requires |
|---|---|---|
| **Demo** | Full pre-built VCF Private Cloud analysis — instant, no key | Nothing |
| **Claude API** | Live generation via `claude-haiku-4-5` | `sk-ant-...` key |
| **OpenAI API** | Live generation via `gpt-4o-mini` | `sk-...` key |

Switch modes with one click — same 3-agent pipeline, same output structure.

---

## Web Tool — Quick Start

### Option 1: Use the live demo
Open [danvzla.github.io/ai-sow-orchestrator](https://danvzla.github.io/ai-sow-orchestrator) in any browser. No setup needed.

### Option 2: Run locally
```bash
# Clone the repo
git clone https://github.com/danvzla/ai-sow-orchestrator.git
cd ai-sow-orchestrator

# Open in browser (no server needed — pure HTML/JS)
open index.html          # macOS
start index.html         # Windows
xdg-open index.html      # Linux
```

### Option 3: Deploy to GitHub Pages
See [Deploy to GitHub Pages](#deploy-to-github-pages) below.

---

## Python CLI — Quick Start

### Install dependencies
```bash
# Core (Claude)
pip install anthropic rich python-docx

# Add this if using OpenAI
pip install openai
```

### Set your API key
```bash
# Claude
export ANTHROPIC_API_KEY=sk-ant-...

# OpenAI
export OPENAI_API_KEY=sk-...
```

### Run
```bash
# Using a notes file — Claude (default)
python sow_generator.py sample_notes/vcf_private_cloud.txt

# Using OpenAI
python sow_generator.py sample_notes/vcf_private_cloud.txt --provider openai

# Interactive mode — prompts for everything
python sow_generator.py --interactive

# Specify engagement context directly
python sow_generator.py notes.txt \
  --type "Telco NFV Transformation" \
  --size "Enterprise ($2M+)" \
  --industry "Telco / CSP"

# Skip Word doc, custom output name
python sow_generator.py notes.txt --no-docx --output my_proposal

# Pass API key directly (not recommended for shared environments)
python sow_generator.py notes.txt --api-key sk-ant-...
```

### Output files
Each run produces:
```
SOW_private_cloud_deployment_vcf_20260617_143022.md      ← Markdown
SOW_private_cloud_deployment_vcf_20260617_143022.docx    ← Word document (charcoal & gold theme)
```

---

## CLI Reference

```
usage: sow_generator.py [-h] [--provider {claude,openai}]
                        [--type TYPE] [--size SIZE] [--industry INDUSTRY]
                        [--output OUTPUT] [--no-docx] [--interactive]
                        [--api-key API_KEY]
                        [notes_file]

Arguments:
  notes_file              Path to discovery notes .txt file

Options:
  --provider {claude,openai}   AI provider (default: claude)
  --type TYPE                  Engagement type (see list below)
  --size SIZE                  Engagement size (see list below)
  --industry INDUSTRY          Customer industry
  --output OUTPUT              Output filename stem (no extension)
  --no-docx                    Skip Word document export
  --interactive                Interactive mode — prompts for all inputs
  --api-key API_KEY            API key (or use env var)
  -h, --help                   Show this help message
```

**Engagement types:**
- Private Cloud Deployment (VCF)
- Network Security Hardening (NSX)
- Telco NFV Transformation
- SASE / SD-WAN Migration
- Cloud Modernization Assessment
- NOC Automation Program
- Data Center Consolidation
- Hybrid Cloud Architecture

**Engagement sizes:** Small (under $150K) · Medium ($150K–$500K) · Large ($500K–$2M) · Enterprise ($2M+)

**Industries:** Financial Services · Healthcare · Telco / CSP · Manufacturing · Technology · Government · Energy · Retail

---

## SOW Structure (14 Sections)

Every generated SOW includes:

1. Engagement Overview
2. Engagement Objective
3. Scope of Work
4. Deliverables *(with effort estimates)*
5. Work Breakdown & Schedule *(phase-by-phase with hours)*
6. Milestones & Acceptance Criteria
7. Assumptions
8. Customer Responsibilities
9. Out of Scope
10. Risk Register
11. Team Structure
12. Open Items
13. Commercial Terms
14. Document Control

---

## Pre-Built Scenarios (Web Tool)

Six scenarios are built into the web tool — load any of them instantly in Demo Mode:

| # | Scenario | Industry |
|---|---|---|
| 01 | VCF Private Cloud Deployment | Financial Services |
| 02 | NSX Security Hardening | Healthcare |
| 03 | Telco NFV Transformation | Telco / CSP |
| 04 | SASE / SD-WAN Migration | Manufacturing |
| 05 | Cloud Modernization Assessment | Technology |
| 06 | NOC Automation Program | Cable MSO |

---

## Deploy to GitHub Pages

### Step 1 — Create the repository
```bash
# On GitHub: create a new repo named exactly:
ai-sow-orchestrator
```

### Step 2 — Push the files
```bash
cd ai-sow-orchestrator
git init
git add .
git commit -m "Initial commit — AI SOW Orchestrator v2"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/ai-sow-orchestrator.git
git push -u origin main
```
> Replace `YOUR_USERNAME` with your GitHub username.

### Step 3 — Enable GitHub Pages
1. Go to your repo on GitHub
2. Click **Settings** → **Pages** (left sidebar)
3. Under **Source**, select **Deploy from a branch**
4. Branch: `main` · Folder: `/ (root)`
5. Click **Save**

### Step 4 — Access your live tool
After 1–2 minutes:
```
https://YOUR_USERNAME.github.io/ai-sow-orchestrator
```

> **Security note:** Browser `localStorage` is not an enterprise secrets-management solution. Use only temporary test credentials in the public demo; Demo Mode requires no key.

---

## Repository Structure

```
ai-sow-orchestrator/
│
├── index.html              ← Web tool (open directly in any browser)
├── sow_generator.py        ← Python CLI (Claude + OpenAI)
├── requirements.txt        ← Python dependencies
├── .gitignore
├── README.md
│
└── sample_notes/           ← Ready-to-use discovery note examples
    ├── vcf_private_cloud.txt
    ├── nsx_security_hardening.txt
    └── telco_nfv_transformation.txt
```

---

## Technical Architecture

### Web tool (index.html)
- Pure HTML · CSS · JavaScript — no framework, no build step
- Three provider modes: Demo / Claude API / OpenAI API
- 4 independently scrollable output tabs (Pipeline · Brief · Validation · SOW)
- API keys stored in browser `localStorage` for test convenience; not recommended for production use
- Works offline in Demo Mode

### Python CLI (sow_generator.py)
- Provider abstraction: same 3 prompts route to Claude or OpenAI
- 4-attempt JSON repair chain handles malformed or truncated LLM responses
- Rich terminal output with progress spinners (degrades gracefully without `rich`)
- Word export (`.docx`) with charcoal & gold styling via `python-docx`

### Agent pipeline
```
Discovery Notes + Context
         │
         ▼
   ┌─────────────────┐
   │   Agent 1       │  ─── JSON brief (workstreams, deliverables, risks, phases)
   │   Discovery     │
   │   Analyst       │
   └────────┬────────┘
            │
            ▼
   ┌─────────────────┐
   │   Agent 2       │  ─── approved / approved_with_warnings / rejected
   │   Scope         │       + quality score 0–100
   │   Validator     │
   └────────┬────────┘
            │ (only if approved)
            ▼
   ┌─────────────────┐
   │   Agent 3       │  ─── 14-section SOW in HTML (web) or Markdown (CLI)
   │   SOW Drafter   │
   └─────────────────┘
```

---

## Requirements

### Web tool
- Any modern browser (Chrome, Firefox, Safari, Edge)
- Internet connection (for CDN fonts and icons)
- API key from [Anthropic](https://console.anthropic.com) or [OpenAI](https://platform.openai.com) *(not needed in Demo Mode)*

### Python CLI
```
Python 3.9+
anthropic >= 0.34.0      # always required
rich >= 13.0.0           # optional, recommended
python-docx >= 1.1.0     # optional, for .docx export
openai >= 1.40.0         # only if using --provider openai
```

---

---

## Architecture and Governance

The solution separates discovery analysis, scope validation, drafting, and human approval. Agent 3 runs only after the validator returns an approved status. The resulting SOW is a draft for commercial, legal, delivery, technical, and customer review.

## Deterministic vs. AI-Generated Outputs

**Deterministic:** pipeline order, validator gate behavior, approval states, required SOW sections, file naming, export behavior, CLI arguments, and provider routing.

**AI-generated:** discovery brief, gap analysis, quality rationale, scope, assumptions, exclusions, deliverables, risks, milestones, and draft contractual language.

## Validation and Quality Controls

Current controls include an independent validation stage, quality scoring, rejection of contradictory scope, blocking of drafting when validation fails, structured output, JSON repair, and manual review.

## Security and Data Handling

> **Security note:** Do not enter production or long-lived API credentials into the public browser demonstration.

Use environment variables for CLI credentials. Do not submit confidential customer notes, pricing, personal data, regulated information, or proprietary designs. Browser local storage is not an enterprise secrets-management solution.

## Testing

Current validation covers provider workflow, gate behavior, structured rendering, CLI arguments, Markdown/DOCX generation, sample scenarios, and graceful optional-dependency behavior. Production use requires unit, integration, schema, clause, prompt-regression, document-diff, and security tests.

## Limitations

Generated SOWs are drafts, not executable contracts. Effort, schedule, pricing, responsibilities, assumptions, and acceptance criteria remain illustrative until independently validated. The tool does not replace discovery, architecture review, estimating, legal review, or customer approval.

## Disclaimer

This project is provided for demonstration and educational purposes and does not constitute legal, contractual, commercial, financial, architecture, or delivery advice.

---

## Author

**Daniel Mazzini**
Principal Architect & Senior TPM · Cloud & Infrastructure · Telco · Applied AI

- LinkedIn: [linkedin.com/in/daniel-mazzini-22059734](https://www.linkedin.com/in/daniel-mazzini-22059734/)
- GitHub: [github.com/danvzla](https://github.com/danvzla)
- Portfolio: [danvzla.github.io/ai-sow-orchestrator](https://danvzla.github.io/ai-sow-orchestrator)

