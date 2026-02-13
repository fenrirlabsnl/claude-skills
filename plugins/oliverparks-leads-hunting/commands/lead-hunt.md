---
description: Full pipeline - scrape jobs from both boards and find hiring managers
argument-hint: "<job titles> <location> [--linkedin] [--full] [--export csv|json|xlsx]"
---

# Lead Hunt

> Run the complete lead generation pipeline: parallel job scraping + LinkedIn enrichment.

## Workflow

### 1. Parse Input

Accept pipeline parameters:
- Job titles (required) - single or comma-separated list
- Location (required)
- Options:
  - `--linkedin` - Also find hiring managers at companies
  - `--full` - Get complete job descriptions
  - `--export` - Export format (csv, json, xlsx)
  - `--batch-size` - Companies per LinkedIn batch (default: 5)

### 2. Phase 1: Job Scraping (Parallel)

Use **indeed-scraper** and **stepstone-scraper** skills simultaneously:
- Both boards scanned in parallel for each job title
- Multiple job titles also processed in parallel

### 3. Phase 2: Merge & Deduplicate

Use **job-filtering** skill logic:
- Match jobs by company + similar title
- Merge data from both sources (keep best of each)
- Flag jobs found on both boards

### 4. Phase 3: LinkedIn Search (if --linkedin)

Use **linkedin-leads** skill:
- Extract unique company list from jobs
- Batch companies (default: 5 per batch)
- Wait 30 seconds between batches
- Match leads to job postings

### 5. Phase 4: Output & Export

Display summary:
- Total jobs, unique companies, duplicates
- LinkedIn leads (if searched)
- Hiring managers vs HR contacts

Export if requested:
- CSV: Flat file with jobs + leads
- JSON: Full structured data
- XLSX: Multi-sheet workbook

## Examples

```
# Basic: scan both job boards
/lead-hunt "SAP FI/CO" "Germany"

# Multiple roles
/lead-hunt "SAP FI/CO, SAP HCM, SAP MM" "Germany"

# Full pipeline with LinkedIn
/lead-hunt "SAP FI/CO" "Germany" --linkedin

# With export
/lead-hunt "Data Engineer" "Berlin" --linkedin --export csv

# Full descriptions
/lead-hunt "Product Manager" "München" --full
```

## Output Format

### Summary View

```
╔══════════════════════════════════════════════════════════════╗
║              LEAD HUNT RESULTS: SAP FI/CO                    ║
╠══════════════════════════════════════════════════════════════╣
║  Jobs Found:        47 (Indeed: 28, Stepstone: 31)           ║
║  Unique Companies:  35                                        ║
║  Duplicates:        12 (found on both boards)                ║
║  LinkedIn Leads:    89                                        ║
║  Hiring Managers:   52                                        ║
║  HR Contacts:       37                                        ║
╚══════════════════════════════════════════════════════════════╝
```

### Top Opportunities Table

```
| # | Company | Role | Location | Salary | Leads |
|---|---------|------|----------|--------|-------|
| 1 | N-ERGIE | SAP FI/CO Berater | Nürnberg | 72-90k | 3 (2 HM) |
| 2 | Bosch | SAP CO Controller | Stuttgart | 68-85k | 2 (1 HM) |
| 3 | Siemens | SAP FI Lead | Erlangen | 80-100k | 4 (3 HM) |
```

## Execution Modes

### Mode 1: Job Scan Only (Default)

```
/lead-hunt "SAP FI/CO" "Germany"
```

- Parallel Indeed + Stepstone scan
- Merge and deduplicate
- Display results
- ~2-3 minutes

### Mode 2: Full Pipeline with LinkedIn

```
/lead-hunt "SAP FI/CO" "Germany" --linkedin
```

- Parallel job scan (Phase 1)
- Merge and extract companies (Phase 2)
- Batched LinkedIn search (Phase 3)
- Combine and rank (Phase 4)
- ~10-15 minutes for 30 companies

### Mode 3: Multi-Role Scan

```
/lead-hunt "SAP FI/CO, SAP HCM, SAP MM" "Germany"
```

- 6 parallel tasks (2 boards × 3 roles)
- Merge by role
- Deduplicate across roles
- ~3-5 minutes

## Rate Limits

### Job Scraping
- Max 6 concurrent scraper tasks
- 2-second stagger between spawns

### LinkedIn
- Default batch size: 5 companies
- 30-second wait between batches
- Max 6 batches per session (30 companies)

## Error Handling

| Issue | Action |
|-------|--------|
| Indeed blocked | Continue with Stepstone |
| Stepstone blocked | Continue with Indeed |
| LinkedIn rate limit | Save partial, stop LinkedIn |
| No jobs found | Report empty, skip LinkedIn |

## Follow-up Commands

After results:
```
"Get full descriptions for the top 5 jobs"
"Show me only jobs with hiring manager leads"
"Filter to remote-friendly positions"
"Find more leads at Siemens and BMW"
"Export results to csv"
```

## Resume Capability

If interrupted:
```
"Resume lead hunt from company #15"
"Continue LinkedIn search for remaining companies"
```

## Daily Workflow Example

```
08:00 - Morning Scan
────────────────────
/lead-hunt "SAP FI/CO, SAP HCM" "Germany"
→ Quick scan of new jobs

09:00 - LinkedIn Enrichment
────────────────────────────
/lead-hunt "SAP FI/CO" "Germany" --linkedin
→ Add hiring manager contacts

10:00 - Export & Outreach
────────────────────────────
"Export results to csv"
→ Import to CRM, begin outreach
```

## Tips

- Use without `--linkedin` for quick daily scans
- Add `--linkedin` when you're ready for deep enrichment
- Multi-role searches are efficient for broad scanning
- Jobs on both boards = active hiring, prioritize these
- Export to CSV for CRM import
