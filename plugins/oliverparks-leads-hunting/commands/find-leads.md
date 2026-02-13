---
description: Find hiring managers on LinkedIn for target companies
argument-hint: "<job title> <company1, company2, ...> [--platform standard|talent]"
---

# Find Leads

> Find decision-makers at target companies for direct outreach.

## Workflow

### 1. Understand the Request

Accept lead search parameters:
- Job title being hired for (required) - used to determine department/function
- Company names (required) - comma-separated list
- Platform (optional): standard (free LinkedIn) or talent (Recruiter/Sales Nav)

### 2. Map Job to Search Profile

The **linkedin-leads** skill maps job titles to relevant decision-maker searches:

| Job Title | Search Profile | Search Terms |
|-----------|---------------|--------------|
| SAP FI/CO | Finance/Controlling | Head of Controlling, CFO, Finance Director |
| SAP HCM | HR | HR Director, CHRO, Head of HR |
| Data Engineer | Data | Head of Data, CDO, Analytics Director |
| Software Engineer | Engineering | VP Engineering, CTO, Engineering Director |

### 3. Execute LinkedIn Search

For each company:
1. Normalize company name (remove legal suffixes)
2. Search LinkedIn People with company filter
3. Extract visible profiles
4. Score by confidence (High/Medium/Low)
5. Flag HR contacts separately from hiring managers

### 4. Report Results

Output for each company:
- Leads found (with hidden profiles noted)
- Name, title, LinkedIn URL
- Confidence score
- HR vs. hiring manager classification

## Examples

```
# Basic search
/find-leads "SAP FI/CO Consultant" "N-ERGIE, Bosch, Siemens"

# Single company
/find-leads "Data Engineer" "Zalando"

# With Recruiter/Sales Navigator
/find-leads "Product Manager" "BMW, Audi" --platform talent
```

## Output Format

```
## Company: N-ERGIE

Found 4 leads (2 hidden profiles skipped):

| Name | Title | Confidence | HR? |
|------|-------|------------|-----|
| Klaus Weber | Head of Controlling | High | No |
| Maria Fischer | Finance Director | High | No |
| Stefan Braun | CFO | High | No |
| Lisa Müller | Talent Acquisition | Low | Yes |
```

## Integration with Job Scraping

Typical workflow:
1. Run `/scrape-jobs` to find companies hiring
2. Extract company names from results
3. Run `/find-leads` with those companies
4. Prioritize outreach to hiring managers

```
# Step 1: Find jobs
/scrape-jobs "SAP FI/CO" "Germany"
→ N-ERGIE, Magni, Gasunie, Bosch, Siemens...

# Step 2: Find leads at those companies
/find-leads "SAP FI/CO" "N-ERGIE, Magni, Gasunie, Bosch, Siemens"
```

## Confidence Scoring

| Level | Meaning | Outreach Priority |
|-------|---------|-------------------|
| High | Title matches function (Head of Controlling for FI/CO) | First |
| Medium | Related department or senior title | Second |
| Low | Generic title or HR contact | Fallback |

## Rate Limits

- Process 1 company at a time
- Wait 2-3 seconds between searches
- Maximum ~25-30 searches per session
- Stop immediately on LinkedIn warnings

## Follow-up Actions

- "Show leads for company X" - Detailed view
- "Find more leads at BMW" - Extend search
- "Export leads to csv" - Download results
- "Get high-confidence leads only" - Filter results

## Tips

- Hiring managers (non-HR) are your primary targets
- HR contacts are useful as backup or for referrals
- High-confidence leads have titles directly matching the job function
- Use `--platform talent` if you have LinkedIn Recruiter or Sales Navigator
