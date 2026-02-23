# LinkedIn Extraction Scripts

## Profile Type Determination

Based on the job title, select the **first matching** profile type:

```javascript
const consultingKeywords = [
  'consulting', 'beratung', 'advisory', 'consultancy'
];
const knownConsultingFirms = [
  'accenture', 'deloitte', 'pwc', 'ey', 'kpmg', 'capgemini', 'wavestone',
  'bearing point', 'bearingpoint', 'ntt data', 'atos', 'cgi', 'infosys',
  'wipro', 'tcs', 'cognizant', 'mckinsey', 'bcg', 'matrix systems'
];

function isConsultingFirm(companyName) {
  const lower = companyName.toLowerCase();
  return consultingKeywords.some(kw => lower.includes(kw))
      || knownConsultingFirms.some(f => lower.includes(f));
}

function determineSearchProfile(jobTitle, companyName) {
  // Consulting firm override — search for Partners, not client-side decision-makers
  if (companyName && isConsultingFirm(companyName)) {
    return {
      type: 'Consulting Leadership',
      searchTerms: ['Partner', 'Managing Director', 'Practice Lead', 'Associate Partner', 'Director']
    };
  }

  const title = jobTitle.toLowerCase();

  const profiles = [
    {
      keywords: ['sap fi', 'sap co', 'fico', 'fi/co', 'controlling'],
      type: 'SAP FICO',
      searchTerms: ['Head of Controlling', 'Finance Director', 'CFO']
    },
    {
      keywords: ['sap hcm', 'successfactors', 'payroll', 'hr module'],
      type: 'SAP HCM',
      searchTerms: ['HR Director', 'Head of HR', 'CHRO', 'Head of People', 'Personalleiter', 'Leitung Personal', 'Head of SAP', 'IT Director']
    },
    {
      keywords: ['sap mm', 'sap sd', 'supply chain', 'logistics', 'procurement'],
      type: 'SAP Supply Chain',
      searchTerms: ['Supply Chain Director', 'Head of Logistics', 'VP Supply Chain']
    },
    {
      keywords: ['sap pp', 'manufacturing', 'plant', 'production'],
      type: 'SAP Manufacturing',
      searchTerms: ['Plant Manager', 'Head of Production', 'VP Manufacturing']
    },
    {
      keywords: ['sap abap', 'sap basis', 'sap technical', 'sap developer'],
      type: 'SAP Technical',
      searchTerms: ['IT Director', 'Head of SAP', 'CIO']
    },
    {
      keywords: ['sap'],
      type: 'SAP Generic',
      searchTerms: ['Head of SAP', 'IT Director', 'CIO']
    },
    {
      keywords: ['engineer', 'developer', 'software', 'backend', 'frontend', 'fullstack'],
      type: 'Engineering',
      searchTerms: ['VP Engineering', 'CTO', 'Engineering Director', 'Head of Engineering']
    },
    {
      keywords: ['data', 'analytics', 'bi', 'business intelligence', 'machine learning'],
      type: 'Data',
      searchTerms: ['Head of Data', 'CDO', 'Analytics Director', 'VP Data']
    },
    {
      keywords: ['marketing', 'brand', 'growth', 'digital marketing'],
      type: 'Marketing',
      searchTerms: ['CMO', 'VP Marketing', 'Head of Marketing', 'Marketing Director']
    },
    {
      keywords: ['sales', 'account', 'business dev', 'revenue'],
      type: 'Sales',
      searchTerms: ['VP Sales', 'CRO', 'Sales Director', 'Head of Sales']
    },
    {
      keywords: ['product manager', 'product owner', 'product lead'],
      type: 'Product',
      searchTerms: ['Head of Product', 'CPO', 'VP Product', 'Product Director']
    },
    {
      keywords: ['design', 'ux', 'ui', 'creative'],
      type: 'Design',
      searchTerms: ['Head of Design', 'Creative Director', 'VP Design']
    },
    {
      keywords: ['operations', 'admin', 'finance', 'accounting'],
      type: 'Operations',
      searchTerms: ['COO', 'Head of Operations', 'Finance Director', 'VP Operations']
    }
  ];

  for (const profile of profiles) {
    if (profile.keywords.some(kw => title.includes(kw))) {
      return profile;
    }
  }

  return {
    type: 'Generic',
    searchTerms: ['HR Director', 'Talent Acquisition', 'Head of HR']
  };
}
```

---

## Confidence Scoring

```javascript
function assignConfidence(profileType, personTitle) {
  const title = personTitle.toLowerCase();

  const highConfidenceMap = {
    'Consulting Leadership': ['partner', 'managing director', 'practice lead', 'associate partner'],
    'SAP FICO': ['controlling', 'finance director', 'cfo', 'financial controller'],
    'SAP HCM': ['hr director', 'chro', 'head of hr', 'head of people', 'personalleiter', 'leitung personal', 'people', 'talent acquisition', 'recruiter', 'head of sap', 'it director'],
    'SAP Supply Chain': ['supply chain', 'logistics', 'procurement'],
    'SAP Manufacturing': ['plant manager', 'production', 'manufacturing'],
    'SAP Technical': ['head of sap', 'it director', 'cio'],
    'Engineering': ['vp engineering', 'cto', 'engineering director'],
    'Data': ['head of data', 'cdo', 'analytics director'],
    'Marketing': ['cmo', 'marketing director', 'head of marketing'],
    'Sales': ['vp sales', 'cro', 'sales director'],
    'Product': ['head of product', 'cpo', 'vp product'],
    'Design': ['head of design', 'creative director'],
    'Operations': ['coo', 'head of operations']
  };

  const keywords = highConfidenceMap[profileType] || [];
  if (keywords.some(kw => title.includes(kw))) return 'high';

  // HR/recruiting titles get medium confidence — they often posted the role
  const hrKeywords = ['recruiter', 'recruiting', 'talent acquisition', 'hr ',
    'human resources', 'people operations', 'hr business partner'];
  if (hrKeywords.some(kw => title.includes(kw))) return 'medium';

  if (title.includes('director') || title.includes('head of') || title.includes('vp ') || title.includes('chief')) {
    return 'medium';
  }

  return 'low';
}
```

---

## Company Name Normalization

Try these variations **in order** until LinkedIn's company dropdown shows a match:

| Priority | Transformation | Example |
|----------|---------------|---------|
| 1 | First 2 words (after removing suffixes) | "Media Saturn" |
| 2 | Without legal suffixes | "Media-Saturn Deutschland" |
| 3 | Original name | "Media-Saturn Deutschland GmbH" |
| 4 | First word only | "Media" |

```javascript
const legalSuffixes = [
  'GmbH', 'SE', 'AG', 'Inc', 'Inc.', 'Ltd', 'Ltd.', 'KG', 'Co.', 'Corp', 'Corp.',
  'LLC', 'B.V.', 'S.A.', 'S.A', 'PLC', 'N.V.', 'e.V.', 'KGaA', 'mbH', 'OHG',
  '& Co.', '& Co', 'Co., KG', 'Deutschland', 'Germany', 'Europe'
];

function normalizeCompanyName(original) {
  // Step 1: Remove legal suffixes
  let cleaned = original;
  for (const suffix of legalSuffixes) {
    const regex = new RegExp(`\\s*${suffix.replace('.', '\\.')}\\s*$`, 'i');
    cleaned = cleaned.replace(regex, '').trim();
  }

  const words = cleaned.split(/[\s-]+/);
  const variations = [];

  // Priority 1: First 2 words (best for LinkedIn dropdown matching)
  if (words.length > 2) {
    variations.push(words.slice(0, 2).join(' '));
  }

  // Priority 2: Without legal suffixes
  if (cleaned !== original) {
    variations.push(cleaned);
  }

  // Priority 3: Original name
  variations.push(original);

  // Priority 4: First word only (last resort)
  if (words.length > 1) {
    variations.push(words[0]);
  }

  return variations;
}
```

---

## Profile Extraction

```javascript
function extractLinkedInProfiles() {
  const profiles = [];
  const hrKeywords = [
    'recruiter', 'recruiting', 'talent acquisition', 'hr ', 'human resources',
    'people operations', 'staffing', 'sourcer', 'employer branding',
    'head of people', 'chief people'
  ];

  // Method 1 (PREFERRED - durable): Find profile links by URL pattern.
  // LinkedIn won't change /in/ URL structure without breaking the internet.
  // This survives CSS class name rotations that break Method 2.
  // Scope to <main> to avoid matching nav links ("Sign in", "Join now", etc.)
  const searchArea = document.querySelector('main') || document;
  const profileLinks = searchArea.querySelectorAll('a[href*="/in/"]');
  if (profileLinks.length > 0) {
    const seen = new Set();
    profileLinks.forEach(link => {
      let url = link.href?.split('?')[0] || '';
      if (!url || seen.has(url)) return;
      seen.add(url);

      // Walk up to find the result card container
      const card = link.closest('[data-view-name="search-entity-result-universal-template"]')
                || link.closest('li')
                || link.parentElement?.parentElement?.parentElement;
      if (!card) return;

      const name = link.innerText?.split('•')[0]?.trim() || '';
      if (!name || name === 'LinkedIn Member' || name.length < 2) return;

      // Title is the next meaningful text block after the name.
      // Skip connection degree lines like "• 1st", "• 2nd", "• 3rd+".
      const allText = card.innerText || '';
      const lines = allText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      const nameIdx = lines.findIndex(l => l.includes(name));
      let title = '';
      if (nameIdx >= 0) {
        for (let i = nameIdx + 1; i < lines.length; i++) {
          const line = lines[i];
          // Skip connection degree indicators and short noise
          if (/^[•·]\s*\d*(st|nd|rd|th)\+?$/i.test(line)) continue;
          if (/^\d+(st|nd|rd|th)\+?$/i.test(line)) continue;
          title = line;
          break;
        }
      }

      const isHR = hrKeywords.some(kw => title.toLowerCase().includes(kw));
      profiles.push({ name, title, url, isHR });
    });

    if (profiles.length > 0) return profiles;
  }

  // Method 2 (FALLBACK - fragile, may break on LinkedIn CSS deploys):
  // These class names rotate periodically. If Method 1 returned results, skip this.
  const cards = document.querySelectorAll('[data-chameleon-result-urn], .reusable-search__result-container');

  cards.forEach(card => {
    const nameEl = card.querySelector('.entity-result__title-text a span[aria-hidden="true"]') ||
                   card.querySelector('.entity-result__title-text a');
    const name = nameEl?.innerText?.split('•')[0]?.trim() || '';

    if (!name || name === 'LinkedIn Member') return;

    const titleEl = card.querySelector('.entity-result__primary-subtitle') ||
                    card.querySelector('.entity-result__summary');
    const title = titleEl?.innerText?.trim() || '';

    const linkEl = card.querySelector('a[href*="/in/"]');
    let url = linkEl?.href || '';
    url = url.split('?')[0];

    const isHR = hrKeywords.some(kw => title.toLowerCase().includes(kw));
    profiles.push({ name, title, url, isHR });
  });

  return profiles;
}
```
