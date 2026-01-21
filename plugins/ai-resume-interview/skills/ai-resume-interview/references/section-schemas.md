# Section Schemas

JSON output formats for each interview section. These match the database schema in `src/types/database.ts`.

---

## Profile Schema

**Table:** `candidate_profile`
**Hook:** `useUpdateProfile()`
**Operation:** Upsert (creates or updates)

```json
{
  "name": "string (required)",
  "email": "string (required)",
  "title": "string (required)",
  "target_titles": ["string array - titles being targeted"],
  "target_company_stages": ["string array - e.g., 'Series B', 'Growth'"],
  "elevator_pitch": "string - 30-second pitch",
  "career_narrative": "string - thread connecting roles",
  "looking_for": "string - what they want in next role",
  "not_looking_for": "string - what they're avoiding",
  "salary_min": "number | null",
  "salary_max": "number | null",
  "availability_status": "'actively_looking' | 'open' | 'not_looking'",
  "availability_date": "string (ISO date) | null",
  "location": "string",
  "remote_preference": "'remote' | 'hybrid' | 'onsite' | 'flexible'",
  "github_url": "string | null",
  "linkedin_url": "string | null",
  "twitter_url": "string | null"
}
```

### Example

```json
{
  "name": "Alex Chen",
  "email": "alex@example.com",
  "title": "Senior Software Engineer",
  "target_titles": ["Staff Engineer", "Engineering Manager"],
  "target_company_stages": ["Series B", "Growth"],
  "elevator_pitch": "I build reliable systems at scale. 8 years turning ambiguous requirements into production systems that don't wake people up at 3am.",
  "career_narrative": "Started as a frontend dev who kept asking 'but why is it slow?' Moved to backend, then infrastructure. Now I bridge the gap between product wants and what systems can deliver.",
  "looking_for": "Technical challenges with real business impact. A team that ships and learns.",
  "not_looking_for": "Feature factories. Roles where engineering is an afterthought. Places where 'move fast' means 'ignore quality'.",
  "salary_min": 200000,
  "salary_max": 280000,
  "availability_status": "open",
  "availability_date": null,
  "location": "San Francisco, CA",
  "remote_preference": "hybrid",
  "github_url": "https://github.com/alexchen",
  "linkedin_url": "https://linkedin.com/in/alexchen",
  "twitter_url": null
}
```

---

## Experience Schema

**Table:** `experiences`
**Hook:** `useCreateExperience()`
**Operation:** Insert (one per role)

```json
{
  "company_name": "string (required)",
  "title": "string (required)",
  "title_progression": "string | null - if promoted during tenure",
  "start_date": "string (ISO date, required)",
  "end_date": "string (ISO date) | null - null if current",
  "is_current": "boolean (required)",
  "bullet_points": ["string array - CV accomplishments"],
  "why_joined": "string | null - real reason for joining",
  "why_left": "string | null - real reason for leaving",
  "actual_contributions": "string | null - beyond the bullet points",
  "proudest_achievement": "string | null - single best thing",
  "would_do_differently": "string | null - hindsight reflection",
  "challenges_faced": "string | null - hardest parts",
  "lessons_learned": "string | null - what the job taught them",
  "manager_would_say": "string | null - manager's honest view",
  "reports_would_say": "string | null - if had reports",
  "quantified_impact": "object | null - metrics and numbers",
  "display_order": "number (required) - for ordering"
}
```

### Example

```json
{
  "company_name": "TechCorp",
  "title": "Senior Software Engineer",
  "title_progression": "Joined as SWE II, promoted to Senior after 14 months",
  "start_date": "2021-03-01",
  "end_date": null,
  "is_current": true,
  "bullet_points": [
    "Led migration to microservices architecture serving 10M+ daily requests",
    "Reduced API latency by 40% through caching and query optimization",
    "Mentored 3 junior engineers"
  ],
  "why_joined": "Engineering culture seemed strong. Friends from previous company had joined and loved it. Also the salary bump didn't hurt.",
  "why_left": null,
  "actual_contributions": "The microservices thing was mostly me pushing for it. Also quietly fixed the deployment pipeline that everyone complained about but no one owned.",
  "proudest_achievement": "Convinced leadership to delay a feature launch by 2 weeks to fix tech debt. Ended up preventing a major outage.",
  "would_do_differently": "Would have documented decisions better. When I went on vacation, three people asked me the same questions.",
  "challenges_faced": "Inherited a codebase with zero tests. Had to convince team that slowing down to add tests would speed us up long-term.",
  "lessons_learned": "The best code is code you don't have to write. Delete before you add.",
  "manager_would_say": "Reliable, opinionated, sometimes too attached to technical perfection. Needs to pick battles better.",
  "reports_would_say": null,
  "quantified_impact": {
    "latency_reduction": "40%",
    "daily_requests": "10M+",
    "engineers_mentored": 3
  },
  "display_order": 1
}
```

---

## Skill Schema

**Table:** `skills`
**Hook:** `useCreateSkill()`
**Operation:** Insert (one per skill)

```json
{
  "skill_name": "string (required)",
  "category": "string (required) - skill type",
  "self_rating": "number 1-10 | null",
  "evidence": "string | null - what proves the rating",
  "honest_notes": "string | null - caveats and context",
  "years_experience": "number | null",
  "last_used": "string (ISO date) | null"
}
```

### Categories

Use these categories for `category` field:
- `Language` - Programming languages (TypeScript, Python, Go)
- `Framework` - Frameworks and libraries (React, Node.js, Django)
- `Tool` - Development tools (Git, Docker, VS Code)
- `Cloud` - Cloud platforms (AWS, GCP, Azure)
- `DevOps` - Infrastructure and ops (Kubernetes, Terraform, CI/CD)
- `API` - API technologies (REST, GraphQL, gRPC)
- `AI` - AI/ML technologies (LLMs, ML frameworks)
- `Product` - Product management skills
- `Research` - User research, market research
- `Analytics` - Data analysis, metrics
- `Leadership` - People management, stakeholder management
- `Process` - Methodologies (Agile, Scrum)
- `Technical` - General technical skills
- `Business` - Business domain knowledge

### Rating Scale

| Rating | Meaning |
|--------|---------|
| 1-4 | Learning/Growth - need help or guidance |
| 5-6 | Moderate - can do the work, might need occasional help |
| 7-8 | Strong - could pass a technical interview |
| 9-10 | Expert - could teach others, deep expertise |

### Example

```json
{
  "skill_name": "TypeScript",
  "category": "Language",
  "self_rating": 8,
  "evidence": "Primary language for 4 years. Built type-safe APIs, complex generic utilities. Contributed to team coding standards.",
  "honest_notes": "Strong in application code. Advanced type gymnastics (conditional types, template literals) I sometimes have to look up.",
  "years_experience": 4,
  "last_used": "2024-01-15"
}
```

```json
{
  "skill_name": "Kubernetes",
  "category": "DevOps",
  "self_rating": 5,
  "evidence": "Can deploy, scale, debug basic issues. Understand pods, services, deployments.",
  "honest_notes": "Know enough to be useful. Complex networking or custom operators? I'm calling the platform team.",
  "years_experience": 2,
  "last_used": "2024-01-10"
}
```

---

## Gap/Weakness Schema

**Table:** `gaps_weaknesses`
**Hook:** `useCreateGap()`
**Operation:** Insert (one per gap)

```json
{
  "gap_type": "string (required) - category of gap",
  "description": "string (required) - what the gap is",
  "why_its_a_gap": "string | null - context on why it matters",
  "interest_in_learning": "boolean (required) - want to improve?"
}
```

### Gap Types

- `Technical` - Missing hard/technical skills
- `Soft Skill` - Communication, presentation, leadership
- `Domain` - Industry or product type experience
- `Experience` - Role or responsibility gaps

### Example

```json
{
  "gap_type": "Technical",
  "description": "Machine learning and AI implementation",
  "why_its_a_gap": "Can discuss ML concepts and work with ML engineers, but can't implement models myself. Fine for product roles, but limits my ability to prototype AI features.",
  "interest_in_learning": true
}
```

```json
{
  "gap_type": "Soft Skill",
  "description": "Large-audience public speaking",
  "why_its_a_gap": "Comfortable in meetings and small groups. Nervous and less effective at conferences or all-hands (100+ people). Actively working on it.",
  "interest_in_learning": true
}
```

```json
{
  "gap_type": "Domain",
  "description": "Enterprise B2B experience",
  "why_its_a_gap": "My background is consumer and SMB. Don't understand enterprise sales cycles, procurement processes, or how to build for IT buyers.",
  "interest_in_learning": false
}
```

---

## FAQ Response Schema

**Table:** `faq_responses`
**Hook:** `useCreateFaq()`
**Operation:** Insert (one per FAQ)

```json
{
  "question": "string (required) - the interview question",
  "answer": "string (required) - honest, prepared answer",
  "is_common_question": "boolean (required) - standard interview Q?"
}
```

### Common Questions to Generate

These should have `is_common_question: true`:
1. "What's your biggest weakness?"
2. "Why are you looking to leave your current role?"
3. "Where do you see yourself in 5 years?"
4. "Why should we hire you?"
5. "What are your salary expectations?"
6. "Tell me about a time you failed."

### Example

```json
{
  "question": "What's your biggest weakness?",
  "answer": "I can be impatient with slow decision-making. When the data is clear and the path is obvious, I want to move. I've learned to slow down and bring people along - explaining my reasoning, asking for their concerns - but it's still something I actively manage.",
  "is_common_question": true
}
```

```json
{
  "question": "Why are you looking to leave?",
  "answer": "The company pivoted from consumer to enterprise. Great opportunity, just not my strength. I stayed through the transition to hand off properly, then started looking for my next consumer-focused role. Happy to connect you with my manager as a reference.",
  "is_common_question": true
}
```

---

## Batch Insert Example

For skills (similar pattern for FAQs, experiences):

```typescript
// In the interview, after collecting all skills
const skills = [
  { skill_name: "TypeScript", category: "Language", self_rating: 8, ... },
  { skill_name: "React", category: "Framework", self_rating: 7, ... },
  // ... more skills
];

// Insert each via the hook
for (const skill of skills) {
  await createSkill.mutateAsync(skill);
}
```

---

## Validation Notes

Before inserting, validate:

1. **Required fields present** - Check all required fields have values
2. **Enum values correct** - availability_status, remote_preference, gap_type
3. **Dates formatted** - ISO 8601 format (YYYY-MM-DD)
4. **Ratings in range** - self_rating between 1-10
5. **Arrays not empty strings** - target_titles, bullet_points should be arrays

### Common Errors

| Error | Fix |
|-------|-----|
| `null value in column "name"` | Ensure required field isn't missing |
| `invalid input value for enum` | Check enum values match exactly |
| `invalid input syntax for type date` | Use ISO format: "2024-01-15" |

---

## Progress Schema

**File:** `interview-data/progress.json`
**Purpose:** Track interview state for resumption and completion tracking

```json
{
  "current_stage": "string - current/next stage to process",
  "last_updated": "string (ISO datetime)",
  "completed": {
    "profile": "boolean",
    "experiences": ["string array - company names completed"],
    "skills": "boolean",
    "gaps": "boolean",
    "faq": "boolean"
  },
  "remaining": {
    "experiences": ["string array - company names not yet processed"],
    "stages": ["string array - stages not yet started"]
  }
}
```

### Stage Values

- `cv_extraction` - Initial CV parsing
- `profile` - Profile quick-fire questions
- `narrative` - Career narrative
- `experiences` - Experience deep dives
- `skills` - Skill assessment
- `gaps` - Gaps and weaknesses
- `faq` - FAQ generation
- `complete` - All stages finished

### Example

```json
{
  "current_stage": "experiences",
  "last_updated": "2024-01-15T10:30:00Z",
  "completed": {
    "profile": true,
    "experiences": ["TechCorp"],
    "skills": false,
    "gaps": false,
    "faq": false
  },
  "remaining": {
    "experiences": ["StartupCo", "Agency Inc"],
    "stages": ["skills", "gaps", "faq"]
  }
}
```

### Resumption Flow

When resuming an interview:

1. Read `progress.json`
2. Report: "Last time we finished [completed items]. Ready to continue with [current_stage]?"
3. Resume from `current_stage` with context from completed files
