# SOW Orchestrator — AgentPM

> Autonomous 3-agent AI pipeline that turns raw customer discovery notes into a professional Statement of Work in under 60 seconds.

**Part of the [AI Portfolio](https://github.com/danvzla) · Daniel Mazzini · Principal Architect & Senior TPM**

---

## What it does

| Agent | Role | Input | Output |
|-------|------|-------|--------|
| **Agent 1 · Discovery Analyst** | Reads · Structures · Extracts | Raw discovery notes | Structured JSON brief |
| **Agent 2 · Scope Validator** | Reviews · Decides · Gates | Agent 1 brief | Approved / Rejected |
| **Agent 3 · SOW Drafter** | Writes · Formats · Delivers | Validated brief | 14-section SOW |

Agent 2 is the key differentiator — it makes an **autonomous go/no-go decision**. If scope contradictions or gaps are found, it rejects the brief and Agent 3 never runs. This is a genuine control loop, not just a pipeline.

---

## Two ways to use it

### Browser (no install)

**[Launch SOW Orchestrator →](https://danvzla.github.io/ai-sow-orchestrator)**

Open in any browser. Enter your [Anthropic API key](https://console.anthropic.com). Select a scenario. Click Generate.

### Python CLI

```bash
# Install dependencies
pip install -r requirements.txt

# Run with a notes file
python sow_generator.py discovery_notes.txt

# Run interactively (prompts for notes and context)
python sow_generator.py --interactive

# Run with all options specified
python sow_generator.py notes.txt \
  --type "Telco NFV Transformation" \
  --size "Enterprise ($2M+)" \
  --industry "Telco / CSP"

# Skip Word document export
python sow_generator.py notes.txt --no-docx
```

**Output files:**
- `SOW_[EngagementType]_[timestamp].md` — Markdown SOW
- `SOW_[EngagementType]_[timestamp].docx` — Formatted Word document

---

## CLI Output

```
╔═══════════════════════════════════════════════════════════════╗
║          SOW Orchestrator — AgentPM AI Portfolio              ║
║          Autonomous 3-Agent Pipeline                          ║
╚═══════════════════════════════════════════════════════════════╝

──────────── Agent 1 · Discovery Analyst ────────────
  Analyzing discovery notes...                    [0:00:12]
  ✓ Brief complete · 9 workstreams · 8 deliverables · 6 risks

──────────── Agent 2 · Scope Validator ─────────────
  Validating scope brief...                       [0:00:08]

  ✓ APPROVED_WITH_WARNINGS · Quality score: 82/100
  Warning  [Dependency]  Ericsson coexistence — integration complexity not captured in deliverables

──────────── Agent 3 · SOW Drafter ─────────────────
  Drafting Statement of Work...                   [0:00:18]
  ✓ SOW drafted · 312 lines · saved to SOW_Telco_NFV.md
  ✓ Word document saved to SOW_Telco_NFV.docx

╔══════════════════════════════════════════════════╗
║  ✓ SOW generation complete                       ║
║  Markdown:  SOW_Telco_NFV.md                     ║
║  Word doc:  SOW_Telco_NFV.docx                   ║
║  Score: 82/100 · APPROVED_WITH_WARNINGS          ║
╚══════════════════════════════════════════════════╝
```

---

## SOW Sections Generated

The SOW Drafter (Agent 3) produces a 14-section document combining professional prose and tables:

1. Engagement Overview
2. Engagement Objective
3. Scope of Work
4. Deliverables *(with quantity and effort hours)*
5. Work Breakdown & Schedule *(with phase subtotals and grand total)*
6. Milestones & Acceptance
7. Assumptions *(with risk if invalid)*
8. Customer Responsibilities
9. Out of Scope
10. Risk Register *(with probability, impact, mitigation)*
11. Team Structure
12. Open Items
13. Commercial Terms
14. Document Control

---

## Validation Gate Examples

**Approved with warnings:**
```
✓ APPROVED (82/100)
⚠ Warning [Dependency]: Ericsson coexistence complexity not in deliverables
  → Agent 3 incorporates warning into risk register automatically
```

**Rejected:**
```
✗ REJECTED (42/100)
  [Contradiction] Zero downtime + 180 VM migration in 9 months
  Fix: Add migration approach and maintenance window schedule
  [Missing] No success criteria for NSX microsegmentation
  Fix: Define acceptance criteria and sign-off process
  → Agent 3 blocked. Fix notes and rerun.
```

---

## Pre-built Demo Scenarios

| # | Scenario | Industry | Size |
|---|----------|----------|------|
| 01 | VCF Private Cloud Deployment | Financial Services | $500K–$2M |
| 02 | NSX Security Hardening | Healthcare | $150K–$500K |
| 03 | Telco NFV Transformation | Tier 2 Carrier | $2M+ |
| 04 | SASE / SD-WAN Migration | Manufacturing | $150K–$500K |
| 05 | Cloud Modernization Assessment | Technology / SaaS | $95K |
| 06 | NOC Automation Program | Cable MSO / Telco | $1.4M |

---

## Requirements

| Package | Version | Purpose |
|---------|---------|---------|
| `anthropic` | ≥ 0.34.0 | Claude API client |
| `rich` | ≥ 13.0.0 | Terminal UI (optional but recommended) |
| `python-docx` | ≥ 1.1.0 | Word document export (optional) |

**API key:** Get a free Anthropic API key at [console.anthropic.com](https://console.anthropic.com)

---

## About this project

This tool automates the discovery-to-proposal workflow I ran hundreds of times across Broadcom, VMware, Dell EMC, and Juniper Networks. The 3-agent architecture separates three distinct cognitive tasks:

- **Analytical extraction** (Agent 1) — reading noise, producing structure
- **Quality gatekeeping** (Agent 2) — catching scope problems before they reach customers
- **Professional writing** (Agent 3) — translating structure into contractually-appropriate language

The autonomous validation gate is what makes this genuinely agentic. Agent 2 can stop the entire pipeline — that's a control loop, not just sequential AI calls.

---

## Author

**Daniel Mazzini** — Principal Architect & Senior TPM  
Cloud & Infrastructure · Pre-Sales & Services Portfolio · Telco · Automation  
Bilingual: English & Spanish

- 📧 danvzla@gmail.com
- 💼 [LinkedIn](https://www.linkedin.com/in/daniel-mazzini-22059734/)
- 🐙 [GitHub](https://github.com/danvzla)
- 🌐 [AgentPM Portfolio](https://danvzla.github.io)
