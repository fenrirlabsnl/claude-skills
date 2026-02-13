---
description: Scrape job listings from Indeed and/or Stepstone Germany
argument-hint: "<job title> <location> [--indeed|--stepstone|--both] [--full]"
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

### 3. Apply Filtering

See the **job-filtering** skill for:
- Agency detection (30+ known agencies)
- Technical role filtering (two-tier logic)
- Deduplication and merging

### 4. Display Results

Use `display_scraped_data` MCP tool to render structured results with:
- Job title and company
- Location and salary
- Remote work status
- Posting time
- Source (Indeed, Stepstone, or both)

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
- Use `--indeed` or `--stepstone` when you know which board has better results for your niche
- `--full` mode is slower but useful for detailed job analysis
- Jobs found on both boards are flagged - these companies are actively hiring
