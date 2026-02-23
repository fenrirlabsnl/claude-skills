# Stepstone Extraction Scripts

## URL Slugification

```javascript
function slugify(text) {
  return text
    .toLowerCase()
    .replace(/[\/\\]/g, '-')       // slashes to hyphens (e.g., FI/CO -> fi-co)
    .replace(/[^a-z0-9äöüß\s-]/g, '') // keep German chars
    .replace(/\s+/g, '-')          // spaces to hyphens
    .replace(/-+/g, '-')           // collapse multiple hyphens
    .replace(/^-|-$/g, '');        // trim leading/trailing hyphens
}

function buildStepstoneURL(jobTitle, location) {
  const titleSlug = slugify(jobTitle);
  const locationSlug = slugify(location);
  return `https://www.stepstone.de/work/${titleSlug}/in-${locationSlug}?radius=30&sort=2&action=facet_selected%3bage%3bage_7&ag=age_1`;
}
```

### Examples

| Input | Slug | Full URL |
|-------|------|----------|
| "SAP FI/CO" + "Germany" | `sap-fi-co` / `germany` | `.../work/sap-fi-co/in-germany?...` |
| "SAP HCM SuccessFactors Consultant Berater" + "Germany" | `sap-hcm-successfactors-consultant-berater` / `germany` | `.../work/sap-hcm-successfactors-consultant-berater/in-germany?...` |
| "Data Engineer" + "Berlin" | `data-engineer` / `berlin` | `.../work/data-engineer/in-berlin?...` |

**Note:** No `q=` query parameter is needed — the slug path is the search query on Stepstone.

---

## Stage 4: Fast Mode Extraction

Extract job data from the **filtered results only**. Stepstone pages contain two sections: the actual search results and a "These jobs might also interest you" recommendation section below. Only extract the filtered results.

**Scoping strategy:** Extract only job cards that appear before any stop boundary — either the "No match yet? There are..." divider or the "These jobs might also interest you" recommendation section.

Apply agency and role filtering using the keyword lists defined in the **job-filtering** skill.

```javascript
// Agency and non-technical keyword lists are defined in the job-filtering skill.
// See skills/job-filtering/SKILL.md for the canonical source.
// Import or inline those lists here at runtime.

function isAgency(company) {
  const lower = company.toLowerCase();
  return AGENCY_KEYWORDS.some(kw => lower.includes(kw));
}

function isNonTechnical(title) {
  const lower = ` ${title.toLowerCase()} `;
  return NON_TECHNICAL_KEYWORDS.some(kw => lower.includes(kw));
}

const jobs = [];
const seen = new Set();

// Stop boundaries — any element containing these marks the end of real results.
// Two types: (1) recommendation sections ("might also interest") appear as headings,
// (2) "No match yet?" dividers appear inline in the results list as a div/p,
//     splitting exact matches from padded/additional results.
const stopKeywords = ['might also interest', 'könnten sie auch interessieren',
  'similar jobs', 'ähnliche jobs', 'recommended for you',
  'no match yet', 'noch kein treffer', 'noch nicht das richtige'];

// Filter out noise phrases that Stepstone injects into card text,
// which can shift line indices and break company name extraction.
const noisePhrases = ['am i a strong match', 'bin ich ein guter match',
  'apply now', 'jetzt bewerben', 'save job', 'job speichern', 'new', 'neu'];

// Find all stop-boundary elements once, then check if a card appears after any of them.
// Scans headings AND generic elements (div, p, span) since the "No match yet?"
// divider is not a heading — it's an inline element within the results list.
function findStopBoundaries() {
  const boundaries = [];
  const candidates = document.querySelectorAll('h2, h3, [role="heading"], div, p, span');
  for (const el of candidates) {
    const text = (el.innerText || '').toLowerCase();
    // Only match if the element's own text is short (< 200 chars) to avoid
    // matching a parent container that happens to contain the keyword deep inside.
    if (text.length < 200 && stopKeywords.some(kw => text.includes(kw))) {
      boundaries.push(el);
    }
  }
  return boundaries;
}

const stopBoundaries = findStopBoundaries();

function isPastStopBoundary(element) {
  for (const boundary of stopBoundaries) {
    // Node.DOCUMENT_POSITION_FOLLOWING = 4
    if (boundary.compareDocumentPosition(element) & Node.DOCUMENT_POSITION_FOLLOWING) {
      return true;
    }
  }
  return false;
}

document
  .querySelectorAll('a[href*="/jobs--"][href*="-inline.html"]')
  .forEach((link) => {
    const href = link.getAttribute("href");
    const jobIdMatch = href.match(/--(\d+)-inline\.html/);
    if (!jobIdMatch) return;

    const jobId = jobIdMatch[1];
    if (seen.has(jobId)) return;
    seen.add(jobId);

    // Resolve the job card container. HTML forbids <a> inside <button>, so
    // closest("button") will never match. Use the job-item ID or <article> instead.
    const card =
      document.getElementById("job-item-" + jobId) ||
      link.closest('article') ||
      link.parentElement?.parentElement?.parentElement?.parentElement
        ?.parentElement;
    if (!card) return;

    // Skip cards past any stop boundary (recommendations or "No match yet?" divider)
    if (isPastStopBoundary(card)) return;

    const title = link.innerText?.trim() || "";

    const lines = card.innerText
      .split("\n")
      .map((l) => l.trim())
      .filter((l) => l && l.length > 1
        && !noisePhrases.some(np => l.toLowerCase() === np));

    let company = "",
      location = "",
      salary = "",
      posted = "",
      workType = "",
      snippet = "";

    // Try company logo alt text first — most reliable source.
    // Stepstone company logos carry the company name in the alt attribute.
    const companyImg = card.querySelector('img[alt]:not([alt=""])');
    const imgCompany = companyImg?.alt?.trim();
    if (imgCompany && imgCompany.length > 1 && imgCompany.length < 100) {
      company = imgCompany;
    }

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (line.includes("€/year")) salary = line;
      else if (
        line.match(/\d+\s*(hours?|Stunden?|days?|Tag)/i) &&
        line.includes("ago")
      )
        posted = line;
      else if (line === "Partially remote" || line === "Fully remote")
        workType = line;
      else if (!company && line === title && lines[i + 1]) {
        company = lines[i + 1];
        location = lines[i + 2] || "";
      } else if (company && !location && line === title && lines[i + 2]) {
        location = lines[i + 2] || "";
      }
    }

    const snippets = lines.filter(
      (l) =>
        l.length > 50 &&
        l !== title &&
        l !== company &&
        !l.includes("€/year") &&
        !l.includes("ago") &&
        l !== "more",
    );
    snippet = snippets[0] || "";

    if (isAgency(company)) return;
    if (isNonTechnical(title)) return;

    jobs.push({
      jobId,  // Raw ID only - URL constructed post-extraction
      href,   // Path portion for URL construction
      title,
      company,
      location,
      workType,
      salary,
      posted,
      description: snippet,
      // url field omitted - constructed post-extraction for MCP compatibility
    });
  });

// Count total cards and how many were skipped by stop boundary
const totalCards = document.querySelectorAll('a[href*="/jobs--"][href*="-inline.html"]').length;
const stopBoundarySkipped = totalCards - seen.size;

JSON.stringify({
  jobs,
  meta: {
    totalCardsOnPage: totalCards,
    exactMatches: jobs.length,
    filteredOut: seen.size - jobs.length,  // agency/non-technical filters
    stopBoundarySkipped: Math.max(0, stopBoundarySkipped)
  }
}, null, 2);
```

---

## Stage 5: Post-Extraction URL Construction

After JavaScript returns the job data, construct URLs in a separate step:

```javascript
// After extraction, map paths to full URLs:
const jobsWithUrls = extractedJobs.map((job) => ({
  ...job,
  url: `https://www.stepstone.de${job.href}`,
}));
```

Or construct URLs when displaying results to the user.

---

## Stage 7: Full Description Extraction

For each employer job, click through and extract full details:

```javascript
function extractFullDescription() {
  // Method 1: job-ad-display container (most reliable)
  const jobAdDisplay = document.querySelector('[class*="job-ad-display"]');
  if (jobAdDisplay && jobAdDisplay.innerText.length > 200) {
    return jobAdDisplay.innerText;
  }

  // Method 2: Find by section keywords
  const keywords = [
    "Das sind wir",
    "Das erwartet Sie",
    "Das bringen Sie mit",
    "Aufgaben",
    "Anforderungen",
    "Benefits",
    "What you",
    "Your responsibilities",
    "Your profile",
  ];
  const allDivs = document.querySelectorAll("div");

  for (const div of allDivs) {
    const text = div.innerText;
    const matchCount = keywords.filter((kw) => text.includes(kw)).length;
    if (matchCount >= 2 && text.length > 500 && text.length < 15000) {
      return text;
    }
  }

  // Method 3: Main content area
  const main =
    document.querySelector("main") || document.querySelector("article");
  if (main && main.innerText.length > 500) {
    return main.innerText;
  }

  return "Full description not found";
}

// Also extract structured sections if available
function extractStructuredDescription() {
  const sections = {};
  const h4s = document.querySelectorAll("h4");

  h4s.forEach((h4) => {
    const title = h4.innerText.trim();
    let content = "";
    let sibling = h4.nextElementSibling;

    while (sibling && sibling.tagName !== "H4") {
      if (sibling.innerText) content += sibling.innerText + "\n";
      sibling = sibling.nextElementSibling;
    }

    if (title && content.trim()) {
      sections[title] = content.trim();
    }
  });

  return sections;
}

JSON.stringify(
  {
    fullText: extractFullDescription(),
    sections: extractStructuredDescription(),
  },
  null,
  2,
);
```

### Full Mode Workflow

For each job in the filtered list, use `browser_snapshot` (not screenshots) for all navigation:

1. **Navigate** to the job URL
2. **Snapshot** - Take accessibility snapshot to verify page loaded
3. **Handle popups** - Press Escape to close any subscription modals
4. **Wait** for page load (check for job-ad-display element in snapshot)
5. **Extract** full description using JavaScript
6. **Return** to search results or next job
