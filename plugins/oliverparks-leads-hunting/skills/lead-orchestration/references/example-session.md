# Example Session

```
You: Hunt leads for SAP FI/CO in Germany with LinkedIn

Claude: Starting lead hunt for SAP FI/CO in Germany with LinkedIn enrichment.

## Phase 1: Job Scraping (Parallel)

[Using indeed-scraper and stepstone-scraper skills...]
✓ Indeed: 28 jobs found
✓ Stepstone: 31 jobs found

## Phase 2: Merge & Deduplicate

- Total raw jobs: 59
- Duplicates found: 12 (same company + similar title)
- Unique jobs: 47
- Unique companies: 35

## Phase 3: LinkedIn Search

Batching 35 companies into 7 batches of 5...

[Batch 1/7] N-ERGIE, Magni, Gasunie, Bosch, Siemens
✓ Found 14 leads (10 hiring managers, 4 HR)

[Batch 2/7] BMW, Allianz, SAP, Deutsche Bank, Henkel
✓ Found 12 leads (8 hiring managers, 4 HR)

[Batch 3/7] ...
...

## Phase 4: Results

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

### Top Opportunities (by lead count)

| # | Company | Role | Location | Salary | Leads |
|---|---------|------|----------|--------|-------|
| 1 | N-ERGIE | SAP FI/CO Berater | Nürnberg | 72-90k | 3 (2 HM) |
| 2 | Bosch | SAP CO Controller | Stuttgart | 68-85k | 2 (1 HM) |
| 3 | Siemens | SAP FI Lead | Erlangen | 80-100k | 4 (3 HM) |
...

Exported to: lead-hunt-sap-fi-co-2026-02-19.xlsx
  4 tabs: Indeed Jobs, Stepstone Jobs, Hiring Managers, HR & Recruiting

Say "show leads for #1" for details.
```
