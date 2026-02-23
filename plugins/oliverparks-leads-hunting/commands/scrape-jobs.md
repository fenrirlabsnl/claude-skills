---
description: Scrape job listings from Indeed and/or Stepstone Germany
argument-hint: "<job title> <location> [--indeed|--stepstone|--both] [--full] [--test]"
allowed-tools: Read, Grep, Glob, mcp__anthropic_chrome__*
---

# Scrape Jobs

> Scrape job listings from German job boards with direct employer filtering.

## Workflow

### 1. Understand the Request

Accept job search parameters:
- Job title or keywords (required)
- Location (required)
- Platform preference: Indeed, Stepstone, or both (default: both)
- Extraction mode: fast (snippets) or full (complete descriptions)

### 2. Route to Appropriate Skill(s)

**If browser automation MCP is connected:**

Based on platform selection:

**If --indeed or no flag specified:**
- Use the **indeed-scraper** skill
- URL-level employer filtering
- Last 24 hours, sorted by date

**If --stepstone or no flag specified:**
- Use the **stepstone-scraper** skill
- Keyword-based agency filtering
- Last 24 hours, sorted by date

**If --both or no flag (default):**
- Run both scrapers in parallel
- Merge results using **job-filtering** skill logic
- Deduplicate jobs found on both boards

**If browser automation MCP is NOT connected:**
Ask the user to:
- Manually search Indeed (de.indeed.com) and/or Stepstone (stepstone.de) and paste job listings
- Provide a CSV export from their job board search
- Share URLs of specific job listings to analyze

### 3. Apply Filtering

See the **job-filtering** skill for:
- Agency detection (30+ known agencies)
- Technical role filtering (exclude-only logic)
- Deduplication and merging

### 4. Display Results

Display structured results to the user with:
- Job title and company
- Location and salary
- Remote work status
- Posting time
- Source (Indeed, Stepstone, or both)

### 5. Output Test Metrics (if --test)

Write `scrape-metrics.csv` with one row per source:

| Column | Description |
|--------|-------------|
| source | "indeed" or "stepstone" |
| raw_jobs | Jobs before any filtering |
| agencies_filtered | Jobs removed by agency detection |
| roles_filtered | Jobs removed by technical role filter |
| final_count | Jobs after all filters |
| duplicates | Jobs also found on other board |

Example:
```
source,raw_jobs,agencies_filtered,roles_filtered,final_count,duplicates
indeed,42,8,3,31,12
stepstone,38,5,2,31,12
```

## Examples

```
# Scan both boards (default)
/scrape-jobs "SAP FI/CO" "Germany"

# Indeed only
/scrape-jobs "Data Engineer" "Berlin" --indeed

# Stepstone only with full descriptions
/scrape-jobs "Product Manager" "München" --stepstone --full

# Explicit both boards
/scrape-jobs "Python Developer" "Hamburg" --both
```

## Output Format

Fast mode (default):
```
| # | Company | Title | Location | Salary | Remote | Source |
|---|---------|-------|----------|--------|--------|--------|
| 1 | N-ERGIE | SAP FI/CO Berater | Nürnberg | 72-90k | Partial | Both |
```

Full mode (--full):
- Includes complete job descriptions
- Structured sections (requirements, responsibilities, benefits)
- Takes longer due to page-by-page extraction

## Follow-up Actions

After results are displayed:
- "Get full description for job #3" - Fetch complete JD
- "Show only remote jobs" - Filter displayed results
- "Find hiring managers for these companies" - Use **find-leads** command
- "Export to csv" - Download results

## Tips

- Default mode scans both boards for comprehensive coverage
- Use `--indeed` or `--stepstone` when one board has better results for the niche
- `--full` mode is slower but useful for detailed job analysis
- Jobs found on both boards are flagged - these companies are actively hiring
