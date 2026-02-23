# Export Format Details

## Table of Contents

- [XLSX Workbook (Default)](#xlsx-workbook-default)
- [generateXLSXData() Implementation](#generatexlsxdata-implementation)
- [CSV Export](#csv-export)
- [JSON Export](#json-export)

---

## XLSX Workbook (Default)

All results are exported as a multi-tab XLSX workbook. Each scraper source gets its own tab. LinkedIn leads are split into two tabs (Hiring Managers and HR & Recruiting) when `--linkedin` is used.

**Tab: Indeed Jobs**
```
Source    | Company                    | Role                      | Location  | Salary         | Remote  | Posted       | Job URL
Indeed    | N-ERGIE Aktiengesellschaft | SAP Inhouse Berater FI/CO | Nürnberg  | 72,000-90,000€ |         | Heute        | https://de.indeed.com/viewjob?jk=abc123
Indeed    | Bosch                      | SAP CO Controller         | Stuttgart |                |         | Vor 12 Std   | https://de.indeed.com/viewjob?jk=def456
```

**Tab: Stepstone Jobs**
```
Source    | Company                    | Role                      | Location  | Salary         | Remote  | Posted       | Job URL
Stepstone | N-ERGIE Aktiengesellschaft | SAP Inhouse Berater FI/CO | Nürnberg  | 72,000-90,000€ | Partial | 5 hours ago  | https://www.stepstone.de/jobs--...
Stepstone | Siemens Energy             | SAP FI Lead               | Erlangen  | 80,000-100,000€| Partial | 3 hours ago  | https://www.stepstone.de/jobs--...
```

**Tab: Hiring Managers** (only with `--linkedin`) — leads where `is_hr === false`
```
Company                    | Job Posted                        | Lead Name       | Lead Title            | Location | Confidence | LinkedIn Profile URL
N-ERGIE Aktiengesellschaft | SAP Inhouse Berater FI/CO (m/w/d) | Max Mustermann  | Head of Controlling   | Nürnberg | high       | https://linkedin.com/in/max-mustermann
Siemens Energy             | SAP FI Lead (m/w/d)               | Stefan Braun    | IT Director           | Erlangen | high       | https://linkedin.com/in/stefan-braun
```

**Tab: HR & Recruiting** (only with `--linkedin`) — leads where `is_hr === true`
```
Company                    | Job Posted                        | Lead Name       | Lead Title              | Location | Confidence | LinkedIn Profile URL
N-ERGIE Aktiengesellschaft | SAP Inhouse Berater FI/CO (m/w/d) | Anna Schmidt    | Talent Acquisition Mgr  | Nürnberg | medium     | https://linkedin.com/in/anna-schmidt
Siemens Energy             | SAP FI Lead (m/w/d)               | Lisa Müller     | HR Business Partner     | Erlangen | medium     | https://linkedin.com/in/lisa-mueller
```

- **Role** in job tabs = the actual scraped job title
- **Job Posted** in lead tabs = the actual scraped job title from the board
- **Lead Title** in lead tabs = the lead's actual LinkedIn title

Filename: `lead-hunt-{job_title}-{YYYY-MM-DD}.xlsx`

---

## generateXLSXData() Implementation

```javascript
function generateXLSXData(indeedJobs, stepstoneJobs, linkedinResults) {
  const indeedSheet = {
    name: 'Indeed Jobs',
    headers: ['Source', 'Company', 'Role', 'Location', 'Salary', 'Remote', 'Posted', 'Job URL'],
    rows: indeedJobs.map(job => [
      'Indeed',
      job.company,
      job.title,
      job.location,
      job.salary || '',
      job.workType || '',
      job.posted || '',
      `https://de.indeed.com/viewjob?jk=${job.jk}`
    ])
  };

  const stepstoneSheet = {
    name: 'Stepstone Jobs',
    headers: ['Source', 'Company', 'Role', 'Location', 'Salary', 'Remote', 'Posted', 'Job URL'],
    rows: stepstoneJobs.map(job => [
      'Stepstone',
      job.company,
      job.title,
      job.location,
      job.salary || '',
      job.workType || '',
      job.posted || '',
      `https://www.stepstone.de${job.href}`
    ])
  };

  const sheets = [indeedSheet, stepstoneSheet];

  if (linkedinResults) {
    const leadHeaders = ['Company', 'Job Posted', 'Lead Name', 'Lead Title', 'Location', 'Confidence', 'LinkedIn Profile URL'];

    const hmSheet = { name: 'Hiring Managers', headers: leadHeaders, rows: [] };
    const hrSheet = { name: 'HR & Recruiting', headers: leadHeaders, rows: [] };

    for (const result of linkedinResults) {
      for (const lead of result.leads) {
        const row = [
          result.company,
          result.title,       // actual scraped job title from board
          lead.name,
          lead.title,
          result.location || '',
          lead.confidence,
          lead.linkedin_url
        ];

        if (lead.is_hr) {
          hrSheet.rows.push(row);
        } else {
          hmSheet.rows.push(row);
        }
      }
    }

    sheets.push(hmSheet);
    sheets.push(hrSheet);
  }

  return sheets;
}
```

Use the **xlsx** skill to write the workbook. Filename format: `lead-hunt-{job_title}-{date}.xlsx`

---

## CSV Export (`--export csv`)

Flat file with all jobs combined. One row per job for job data.

```csv
Source,Company,Role,Location,Salary,Remote,Posted,Job URL
Indeed,N-ERGIE Aktiengesellschaft,SAP Inhouse Berater FI/CO,Nürnberg,"72,000-90,000€",,Heute,https://de.indeed.com/viewjob?jk=abc123
Stepstone,N-ERGIE Aktiengesellschaft,SAP Inhouse Berater FI/CO,Nürnberg,"72,000-90,000€",Partial,5 hours ago,https://www.stepstone.de/jobs--...
```

If `--linkedin`, two additional CSVs are generated:

**Hiring managers:**
```csv
Company,Job Posted,Lead Name,Lead Title,Location,Confidence,LinkedIn Profile URL
N-ERGIE Aktiengesellschaft,SAP Inhouse Berater FI/CO (m/w/d),Max Mustermann,Head of Controlling,Nürnberg,high,https://linkedin.com/in/max-mustermann
```

**HR & Recruiting:**
```csv
Company,Job Posted,Lead Name,Lead Title,Location,Confidence,LinkedIn Profile URL
N-ERGIE Aktiengesellschaft,SAP Inhouse Berater FI/CO (m/w/d),Anna Schmidt,Talent Acquisition Mgr,Nürnberg,medium,https://linkedin.com/in/anna-schmidt
```

Filenames:
- `lead-hunt-{job_title}-{YYYY-MM-DD}-jobs.csv`
- `lead-hunt-{job_title}-{YYYY-MM-DD}-hiring-managers.csv`
- `lead-hunt-{job_title}-{YYYY-MM-DD}-hr-contacts.csv`

---

## JSON Export (`--export json`)

Full structured data as shown in the Output Format section of SKILL.md.

Filename: `lead-hunt-{job_title}-{YYYY-MM-DD}.json`
