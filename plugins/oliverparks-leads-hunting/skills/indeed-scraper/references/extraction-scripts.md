# Indeed Extraction Scripts

## Stage 4: Fast Mode Extraction

Extract job data from Indeed search results. Apply agency and role filtering using the keyword lists defined in the **job-filtering** skill.

```javascript
// Agency and non-technical keyword lists are defined in the job-filtering skill.
// See skills/job-filtering/SKILL.md for the canonical source.
// Import or inline those lists here at runtime.

function isAgency(company) {
  const lower = company.toLowerCase();
  return AGENCY_KEYWORDS.some(kw => lower.includes(kw));
}

function isNonTechnical(title) {
  const lower = title.toLowerCase();
  return NON_TECHNICAL_KEYWORDS.some(kw => lower.includes(kw));
}

const jobs = [];
const seen = new Set();

document.querySelectorAll("[data-jk]").forEach((card) => {
  const jk = card.getAttribute("data-jk");
  if (!jk || seen.has(jk)) return;
  seen.add(jk);

  const box = card.closest('[class*="cardOutline"]') || card;

  const title = box.querySelector("h2")?.innerText?.trim() || "";
  const company =
    box.querySelector('[data-testid="company-name"]')
      ?.innerText?.split("\n")[0]?.trim() || "";

  if (isAgency(company)) return;
  if (isNonTechnical(title)) return;

  const location =
    box.querySelector('[data-testid="text-location"]')?.innerText?.trim() || "";

  const salaryEl = box.querySelector(
    '[class*="salary"], [data-testid="attribute_snippet_testid"]'
  );
  const salary = salaryEl?.innerText?.trim() || "";

  const postedEl = box.querySelector(
    '[class*="date"], [data-testid="myJobsStateDate"]'
  );
  const posted = postedEl?.innerText?.trim() || "";

  const snippetEl = box.querySelector(
    '[class*="job-snippet"], .jobsearch-JobComponent-description'
  );
  let snippet = "";
  if (snippetEl) {
    snippet = snippetEl.innerText?.trim() || "";
  } else {
    const allText = box.innerText || "";
    const lines = allText.split("\n").filter(
      (l) =>
        l.trim() &&
        l.trim() !== title &&
        l.trim() !== company &&
        l.trim() !== location &&
        !l.includes("Heute") &&
        !l.includes("Vor") &&
        l.length > 30
    );
    snippet = lines[0] || "";
  }

  jobs.push({
    jk,
    title,
    company,
    location,
    salary,
    posted,
    description: snippet,
  });
});

JSON.stringify(jobs, null, 2);
```

---

## Stage 7: Full Description Extraction

For each job, click through and extract full details:

```javascript
function extractFullDescription() {
  const descEl =
    document.querySelector("#jobDescriptionText") ||
    document.querySelector('[class*="jobDescriptionText"]') ||
    document.querySelector('[id*="jobDescription"]');
  if (descEl) return descEl.innerText?.trim() || "";

  const fallback = document.querySelector(".jobsearch-JobComponent-description");
  if (fallback) return fallback.innerText?.trim() || "";

  return "Full description not found";
}

function extractJobDetails() {
  return {
    title:
      document.querySelector('[class*="JobInfoHeader"] h2')
        ?.innerText?.replace("- job post", "").trim() || "",
    company:
      document.querySelector('[data-testid="inlineHeader-companyName"]')
        ?.innerText?.trim() || "",
    location:
      document.querySelector('[data-testid="inlineHeader-companyLocation"]')
        ?.innerText?.trim() || "",
    fullDescription: extractFullDescription(),
    contractType:
      document.querySelector('[data-testid="jobAttribute"]')
        ?.innerText?.trim() || "",
  };
}

JSON.stringify(extractJobDetails(), null, 2);
```
