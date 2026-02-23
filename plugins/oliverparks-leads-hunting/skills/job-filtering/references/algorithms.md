# Job Filtering Algorithms

## Agency Detection Function

```javascript
function isAgency(companyName) {
  const lower = companyName.toLowerCase();
  return agencyKeywords.some(kw => lower.includes(kw));
}
```

---

## Technical Role Detection Function

```javascript
function isTechnicalRole(jobTitle) {
  // Pad with spaces so space-bounded keywords (e.g., ' user ') match at start/end
  const lower = ` ${jobTitle.toLowerCase()} `;

  // Exclude non-technical roles
  if (nonTechnicalRoleKeywords.some(kw => lower.includes(kw))) {
    return false;
  }

  // Include everything else - search query provides relevance
  return true;
}
```

---

## Job Deduplication

### Normalization Functions

```javascript
function normalizeCompanyName(company) {
  return company.toLowerCase()
    .replace(/\b(gmbh|se|ag|inc|ltd|kg|co\.|corp|llc|b\.v\.|plc)\b/gi, '')
    .replace(/\b(deutschland|germany|europe)\b/gi, '')
    .replace(/[^a-z0-9]/g, '');
}

function normalizeJobTitle(title) {
  return title.toLowerCase()
    .replace(/\(m\/w\/d\)|\(w\/m\/d\)|\(all genders\)|\(f\/m\/d\)/gi, '')
    .replace(/[^a-z0-9]/g, '')
    .trim();
}

function createJobKey(company, title) {
  return `${normalizeCompanyName(company)}|${normalizeJobTitle(title)}`;
}
```

### Merge Strategy

When the same job appears on multiple boards:

```javascript
function mergeJobListings(existing, newJob) {
  // Keep both source URLs
  existing.sources.push(newJob.source);
  existing.urls[newJob.source] = newJob.url;

  // Prefer longer description
  if (newJob.description?.length > existing.description?.length) {
    existing.description = newJob.description;
  }

  // Add salary if missing (Stepstone usually has salary)
  if (!existing.salary && newJob.salary) {
    existing.salary = newJob.salary;
  }

  // Add work type if missing (Stepstone has remote info)
  if (!existing.workType && newJob.workType) {
    existing.workType = newJob.workType;
  }

  return existing;
}
```

---

## Job Quality Scoring

```javascript
function scoreJob(job) {
  let score = 0;

  // Salary factors
  if (job.salary) {
    score += 2;
    const salaryNum = parseInt(job.salary.replace(/\D/g, ''));
    if (salaryNum >= 80000) score += 1;
  }

  // Remote factors
  if (job.workType) {
    score += 1;
    if (job.workType.toLowerCase().includes('fully')) score += 1;
  }

  // Employer type (direct employer, not agency)
  if (!job.isAgency) score += 3;

  // Freshness
  if (job.posted?.toLowerCase().includes('heute') ||
      job.posted?.toLowerCase().includes('today') ||
      job.posted?.includes('1 hour') ||
      job.posted?.includes('1 Stunde')) {
    score += 2;
  }

  // Description quality
  if (job.fullDescription || job.description?.length > 200) {
    score += 1;
  }

  // Multi-source bonus
  if (job.sources?.length > 1) {
    score += 1;
  }

  return score;
}
```

---

## Company Name Normalization (Space-Preserving)

For matching company names across sources while preserving readability:

```javascript
const legalSuffixes = [
  'GmbH', 'SE', 'AG', 'Inc', 'Inc.', 'Ltd', 'Ltd.', 'KG', 'Co.', 'Corp', 'Corp.',
  'LLC', 'B.V.', 'S.A.', 'S.A', 'PLC', 'N.V.', 'e.V.', 'KGaA', 'mbH', 'OHG',
  '& Co.', '& Co', 'Co., KG', 'Deutschland', 'Germany', 'Europe'
];

function normalizeForMatching(companyName) {
  let cleaned = companyName;

  // Remove legal suffixes
  for (const suffix of legalSuffixes) {
    const regex = new RegExp(`\\s*${suffix.replace('.', '\\.')}\\s*$`, 'i');
    cleaned = cleaned.replace(regex, '').trim();
  }

  // Normalize whitespace and case
  return cleaned.toLowerCase().replace(/\s+/g, ' ').trim();
}
```

---

## Company Name Disambiguation for LinkedIn

```javascript
const ambiguousWords = [
  'group', 'solutions', 'global', 'systems', 'services',
  'technologies', 'digital', 'united', 'prime', 'core', 'one',
  'progressive', 'motive', 'unity', 'matrix', 'apex', 'summit'
];

function needsDisambiguation(companyName) {
  // Use normalizeForMatching (preserves spaces) — not normalizeCompanyName (strips them)
  const normalized = normalizeForMatching(companyName);
  const words = normalized.split(/\s+/);
  if (words.length <= 2) {
    return ambiguousWords.some(aw => words.some(w => w === aw));
  }
  return false;
}
```

---

## Usage in Scraper Skills

```javascript
// During extraction loop
for (const card of jobCards) {
  const company = extractCompanyName(card);
  const title = extractJobTitle(card);

  // Apply filters
  if (isAgency(company)) continue;
  if (!isTechnicalRole(title)) continue;

  // Add to results
  jobs.push({ company, title, ... });
}
```

## Usage in Lead Orchestration

```javascript
// After collecting from both sources
const allJobs = [...indeedJobs, ...stepstoneJobs];
const deduped = new Map();

for (const job of allJobs) {
  const key = createJobKey(job.company, job.title);

  if (deduped.has(key)) {
    mergeJobListings(deduped.get(key), job);
  } else {
    deduped.set(key, { ...job, sources: [job.source], urls: { [job.source]: job.url } });
  }
}

// Score and sort
const scored = Array.from(deduped.values())
  .map(job => ({ ...job, score: scoreJob(job) }))
  .sort((a, b) => b.score - a.score);
```
