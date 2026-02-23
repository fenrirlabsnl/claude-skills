---
description: Full pipeline - scrape jobs from both boards and find hiring managers
argument-hint: "<job titles> <location> [--linkedin] [--full] [--export csv|json] [--test]"
allowed-tools: Read, Grep, Glob, mcp__anthropic_chrome__*
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
  - `--export` - Export format: csv or json (default: xlsx)
  - `--batch-size` - Companies per LinkedIn batch (default: 5)

### 2. Phase 1: Job Scraping (Parallel)

**If browser automation MCP is connected:**

Use **indeed-scraper** and **stepstone-scraper** skills simultaneously:
- Both boards scanned in parallel for each job title
- Multiple job titles also processed in parallel

**If browser automation MCP is NOT connected:**
Ask the user to:
- Manually search Indeed (de.indeed.com) and Stepstone (stepstone.de) for the job title and location
- Paste job listings or provide a CSV/Excel export from their searches
- Share URLs of job listings to analyze
Then continue with Phase 2 using the provided data.

### 3. Phase 2: Merge & Deduplicate

Use **job-filtering** skill logic:
- Match jobs by company + similar title
- Merge data from both sources (keep best of each)
- Flag jobs found on both boards

### 4. Phase 3: LinkedIn Search (if --linkedin)

**If browser automation MCP is connected:**

Use **linkedin-leads** skill:
- Extract unique company list from jobs
- Batch companies (default: 5 per batch)
- Wait 30 seconds between batches
- Match leads to job postings

**If browser automation MCP is NOT connected:**
Ask the user to:
- Manually search LinkedIn People for decision-makers at the discovered companies
- Paste profile names, titles, and LinkedIn URLs found
- Provide contacts from CRM or other sources
Then continue with Phase 4 using the provided data.

### 5. Phase 4: Output & Export

Display summary:
- Total jobs, unique companies, duplicates
- LinkedIn leads (if searched)
- Hiring managers vs HR contacts

Export results (default: XLSX):
- **XLSX** (default): Multi-tab workbook
  - **Indeed Jobs** tab: Source | Company | Role | Location | Salary | Remote | Posted | Job URL
  - **Stepstone Jobs** tab: Source | Company | Role | Location | Salary | Remote | Posted | Job URL
  - **Hiring Managers** tab (if --linkedin): Company | Job Posted | Lead Name | Lead Title | Location | Confidence | LinkedIn Profile URL
  - **HR & Recruiting** tab (if --linkedin): Company | Job Posted | Lead Name | Lead Title | Location | Confidence | LinkedIn Profile URL
- **CSV** (`--export csv`): Flat file with all jobs combined, one row per job/lead
- **JSON** (`--export json`): Full structured data with nested results

### 6. Output Test Metrics (if --test)

Write 3 CSV files covering each pipeline stage:

1. **`scrape-metrics.csv`** — Per-source job scraping stats (same as /scrape-jobs --test)
2. **`merge-metrics.csv`** — Deduplication results:

| Column | Description |
|--------|-------------|
| total_raw | Sum of jobs from all sources |
| indeed_only | Jobs found only on Indeed |
| stepstone_only | Jobs found only on Stepstone |
| found_on_both | Duplicate jobs |
| unique_after_merge | Final deduplicated count |
| unique_companies | Distinct companies |

Example:
```
total_raw,indeed_only,stepstone_only,found_on_both,unique_after_merge,unique_companies
80,19,19,12,47,35
```

3. **`leads-metrics.csv`** — Per-company LinkedIn stats (same as /find-leads --test, only if --linkedin)

## Examples

```
# Basic: scan both job boards
/lead-hunt "SAP FI/CO" "Germany"

# Multiple roles
/lead-hunt "SAP FI/CO, SAP HCM, SAP MM" "Germany"

# Full pipeline with LinkedIn
/lead-hunt "SAP FI/CO" "Germany" --linkedin

# Export as CSV instead of XLSX
/lead-hunt "Data Engineer" "Berlin" --linkedin --export csv

# Full descriptions
/lead-hunt "Product Manager" "München" --full
```

## Output Format

### Summary View

```
+------------------------------------------------------------+
|              LEAD HUNT RESULTS: SAP FI/CO                  |
+------------------------------------------------------------+
|  Jobs Found:        47 (Indeed: 28, Stepstone: 31)         |
|  Unique Companies:  35                                     |
|  Duplicates:        12 (found on both boards)              |
|  LinkedIn Leads:    89                                     |
|  Hiring Managers:   52                                     |
|  HR Contacts:       37                                     |
+------------------------------------------------------------+
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

- 6 parallel tasks (2 boards x 3 roles)
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
--------------------
/lead-hunt "SAP FI/CO, SAP HCM" "Germany"
-> Quick scan of new jobs

09:00 - LinkedIn Enrichment
----------------------------
/lead-hunt "SAP FI/CO" "Germany" --linkedin
-> Add hiring manager contacts

10:00 - Review & Outreach
----------------------------
-> Open lead-hunt-sap-fi-co-2026-02-19.xlsx
-> Import Hiring Managers + HR & Recruiting tabs to CRM, begin outreach
```

## Tips

- Use without `--linkedin` for quick daily scans
- Add `--linkedin` for deep enrichment
- Multi-role searches are efficient for broad scanning
- Jobs on both boards = active hiring, prioritize these
- Results are always exported as a multi-tab XLSX workbook
