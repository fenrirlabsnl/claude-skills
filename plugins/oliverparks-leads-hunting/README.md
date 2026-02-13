# Oliver Parks Leads Hunting

A lead generation plugin for recruiters. Scrapes job listings from Indeed and Stepstone Germany, finds hiring managers on LinkedIn, and orchestrates the full pipeline for maximum efficiency.

## Installation

```
claude plugins add oliverparks-leads-hunting
```

## What It Does

This plugin gives you an AI-powered recruiting assistant that can:

- **Scrape Job Boards** - Parallel scanning of Indeed and Stepstone Germany with direct employer filtering
- **Find Hiring Managers** - LinkedIn profile search with function-based targeting (not just HR gatekeepers)
- **Orchestrate Pipelines** - Full lead generation workflow from job discovery to decision-maker identification
- **Smart Filtering** - Agency detection, technical role filtering, and intelligent deduplication
- **Export Results** - CSV, JSON, or XLSX output for CRM import

## Commands

| Command | What It Does |
|---|---|
| `/scrape-jobs` | Scrape job listings from Indeed and/or Stepstone Germany |
| `/find-leads` | Find hiring managers on LinkedIn for target companies |
| `/lead-hunt` | Full pipeline: parallel job scraping + LinkedIn enrichment |

## Skills

| Skill | What It Covers |
|---|---|
| `indeed-scraper` | Indeed Germany extraction, employer-only filtering, fast/full modes |
| `stepstone-scraper` | Stepstone Germany extraction, agency filtering, salary/remote data |
| `linkedin-leads` | LinkedIn people search, function mapping, confidence scoring |
| `lead-orchestration` | Parallel execution, deduplication, batched LinkedIn, export |
| `job-filtering` | Agency keywords, technical role detection, normalization |

## Example Workflows

### Quick Job Scan

```
You: /scrape-jobs
Claude: What role and location are you searching for?
You: SAP FI/CO in Germany
Claude: [Scans both Indeed and Stepstone in parallel]
Claude: Found 47 jobs from direct employers (12 on both boards)
```

### Full Lead Hunt

```
You: /lead-hunt
Claude: What role and location? Include --linkedin for hiring manager search
You: "SAP FI/CO" "Germany" --linkedin
Claude: [Phase 1: Parallel job scraping]
Claude: [Phase 2: Merge and deduplicate - 35 unique companies]
Claude: [Phase 3: LinkedIn search in batches of 5]
Claude: [Phase 4: Results with 52 hiring managers, 37 HR contacts]
```

### Find Hiring Managers

```
You: /find-leads
Claude: What role and which companies?
You: "SAP FI/CO Consultant" "N-ERGIE, Bosch, Siemens"
Claude: [Searches LinkedIn for Head of Controlling, Finance Director, CFO]
Claude: Found 12 leads across 3 companies (8 hiring managers, 4 HR)
```

## Key Features

### Direct Employer Filtering

Both scrapers filter out recruitment agencies:
- Indeed: URL-level employer-only parameter
- Stepstone: Keyword matching against 30+ known agencies

### Technical Role Detection

Two-tier filtering ensures relevant roles:
1. Exclude: Sachbearbeiter, Buchhalter, Praktikant (end-users)
2. Include: Consultant, Berater, Engineer, Developer (implementers)

### Function-Based LinkedIn Search

Maps job titles to decision-maker searches:
- SAP FI/CO → Head of Controlling, CFO, Finance Director
- Data Engineer → Head of Data, CDO, Analytics Director
- Product Manager → Head of Product, CPO, VP Product

### Smart Deduplication

Jobs found on both boards are:
- Merged (keeping best data from each source)
- Flagged (indicates active hiring)
- Prioritized (companies posting widely = urgent need)

## Data Sources

Connect browser automation tools for the best experience:

**Included MCP connections:**
- Playwright (plugin-playwright) for Indeed and Stepstone scraping
- Claude in Chrome for LinkedIn search

**Alternative options:**
- See [CONNECTORS.md](CONNECTORS.md) for other browser automation tools

## Security

All scraped content is treated as untrusted data:
- Job descriptions may contain injection attempts
- Profile data is user-provided and unverified
- Rate limits are respected to protect accounts
- LinkedIn extraction only - no automated outreach
