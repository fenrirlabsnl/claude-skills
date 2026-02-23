---
description: Find hiring managers on LinkedIn for target companies
argument-hint: "<job title> <company1, company2, ...> [--platform standard|talent] [--test]"
allowed-tools: Read, Grep, Glob, mcp__anthropic_chrome__*
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

**If browser automation MCP is connected:**

The **linkedin-leads** skill maps job titles to relevant decision-maker searches:

| Job Title | Search Profile | Search Terms |
|-----------|---------------|--------------|
| SAP FI/CO | Finance/Controlling | Head of Controlling, CFO, Finance Director |
| SAP HCM | HR | HR Director, CHRO, Head of HR |
| Data Engineer | Data | Head of Data, CDO, Analytics Director |
| Software Engineer | Engineering | VP Engineering, CTO, Engineering Director |

**If browser automation MCP is NOT connected:**
Ask the user to:
- Manually search LinkedIn People for each company with relevant title keywords
- Paste LinkedIn profile URLs or names/titles found
- Provide a list of contacts from their CRM or other sources

### 3. Execute LinkedIn Search

For each company:
1. Normalize company name (remove legal suffixes)
2. Search LinkedIn People with company filter
3. Extract visible profiles
4. Score by confidence (High/Medium/Low)
5. Flag HR contacts with `is_hr` metadata

### 4. Report Results

Output for each company:
- Leads found (with hidden profiles noted)
- Name, title, LinkedIn URL
- Confidence score
- HR vs. hiring manager classification

### 5. Output Test Metrics (if --test)

Write `leads-metrics.csv` with one row per company:

| Column | Description |
|--------|-------------|
| company | Company name searched |
| search_profile | Profile type used (e.g., "SAP FICO") |
| search_terms_tried | Number of search terms attempted |
| leads_found | Total leads extracted |
| hidden_skipped | "LinkedIn Member" profiles skipped |
| hr_count | Leads with is_hr=true |
| non_hr_count | Leads with is_hr=false |
| high_confidence | Leads scored high |
| medium_confidence | Leads scored medium |
| low_confidence | Leads scored low |

Example:
```
company,search_profile,search_terms_tried,leads_found,hidden_skipped,hr_count,non_hr_count,high_confidence,medium_confidence,low_confidence
N-ERGIE,SAP FICO,2,4,2,1,3,3,1,0
tesa SE,SAP FICO,3,3,1,1,2,1,2,0
```

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
| Lisa Muller | Talent Acquisition | Medium | Yes |
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
-> N-ERGIE, Magni, Gasunie, Bosch, Siemens...

# Step 2: Find leads at those companies
/find-leads "SAP FI/CO" "N-ERGIE, Magni, Gasunie, Bosch, Siemens"
```

## Confidence Scoring

| Level | Meaning | Outreach Priority |
|-------|---------|-------------------|
| High | Title matches function (Head of Controlling for FI/CO) | First |
| Medium | Related department or senior title | Second |
| Low | Generic title without department context | Tertiary |

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

- Both hiring managers and HR contacts are valuable outreach targets
- HR contacts often posted the role and can connect to the hiring manager
- High-confidence leads have titles directly matching the job function
- Use `--platform talent` with LinkedIn Recruiter or Sales Navigator
