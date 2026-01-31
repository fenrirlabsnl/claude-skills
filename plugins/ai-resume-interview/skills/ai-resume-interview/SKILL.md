---
name: ai-resume-interview
category: Career
description: >
  Conversational interview skill to populate the AI Resume with brutally honest, nuanced career data.
  Optimized for voice input via Wispr. Use when user says "interview me", "populate my resume",
  "let's do the interview", "/ai-resume-interview", or "/interview". Extracts CV data first, then
  conducts a structured interview covering: profile, career narrative, experience deep-dives,
  honest skill self-assessment, gaps/weaknesses, and FAQ generation. Saves data to JSON files in interview-data/, with optional Supabase publish at the end.
---

# AI Resume Interview Skill

> Conversational interview to populate your AI Resume with brutally honest, nuanced career data. Optimized for voice input via Wispr.

## Philosophy

This skill exists because resume data is shallow. Traditional job applications capture _what_ you did, not _why_ you made decisions, _what_ you learned, or _how_ you honestly assess your abilities. The AI Resume needs depth that only comes from conversation.

**Core principles:**

- **Brutal honesty over polish** - This is for the AI to represent you accurately
- **Voice-first questions** - Under 20 words, one question at a time
- **Progressive depth** - Start with CV extraction, then dig deeper
- **Capability-based output** - Try file save first, fall back to JSON export if unavailable. Works across CLI and Desktop without platform detection.
- **Output** - Varied sentence length, humanized. Reference how the person is entering their answers and match their tone and phrasing. MUST not use em-dash in outputs.

## Workflow Overview

```
Stage 0: CV Foundation  →  Deep extraction from uploaded CV
Stage 1: Profile        →  Quick-fire identity questions
Stage 2: Narrative      →  Career story and goals
Stage 3: Experiences    →  Per-role deep dive (recent first)
Stage 4: Skills         →  Honest self-assessment with evidence
Stage 5: Gaps           →  Weaknesses and growth areas
Stage 6: FAQ            →  Pre-generate interview answers
```

---

## Stage 0: CV Foundation (Deep Extraction)

Before asking anything, extract everything possible from the user's CV.

### Input Options

Ask: **"Got a CV to start from? Paste it, give me a file path, or say 'skip' to start fresh."**

### Input Handling

- **File path** (CLI): Read file directly via Read tool
- **Pasted text** (CLI/Desktop): Parse inline content
- **Attached file** (Desktop): Process attachment content
- **Skip**: Start from scratch with questions

No platform detection needed - accept whatever format the user provides.

### Extraction Targets

Parse and organize:

- **Profile:** name, email, current title, location, LinkedIn/GitHub/Twitter URLs
- **Experiences:** company name, title, dates, all bullet points exactly as written
- **Skills:** every technology, tool, framework mentioned (with context about where used)
- **Education:** degrees, certifications, courses

### Verification

Present extracted data in organized sections:

```
Here's what I found:

**Profile**
- Name: [extracted]
- Title: [extracted]
- Location: [extracted]

**Experience** (3 roles)
1. [Company] - [Title] (dates)
   • [bullet points]
...

**Skills detected:** TypeScript, React, Node.js, PostgreSQL, AWS...

Anything wrong or missing before we continue?
```

Wait for confirmation or corrections. This verified data becomes context for all subsequent questions (e.g., "You worked at Fintech Co for 3 years...").

---

## Stage 1: Profile Quick-Fire

Fast questions to establish identity. One question, one answer.

### Required Fields

| Field                 | Question                                                               |
| --------------------- | ---------------------------------------------------------------------- |
| `name`                | (Usually from CV) "Is [name] correct, or do you go by something else?" |
| `title`               | "What title actually describes what you do?"                           |
| `elevator_pitch`      | "Thirty-second pitch. Go."                                             |
| `location`            | (Usually from CV) "Still based in [location]?"                         |
| `remote_preference`   | "Remote, hybrid, or office. What works for you?"                       |
| `availability_status` | "Actively looking, passively open, or not looking?"                    |
| `looking_for`         | "What are you optimizing for in your next role?"                       |
| `not_looking_for`     | "What are you actively avoiding?"                                      |

### Optional Fields (ask if relevant)

| Field                       | Question                           |
| --------------------------- | ---------------------------------- |
| `target_titles`             | "What titles are you targeting?"   |
| `target_company_stages`     | "Startup, growth, or big company?" |
| `salary_min` / `salary_max` | "What's your target comp range?"   |
| `availability_date`         | "When could you start?"            |
| `github_url`                | (If developer) "GitHub handle?"    |
| `linkedin_url`              | (Usually from CV)                  |

### Output Schema

See `references/section-schemas.md` for the exact JSON format.

### Save Point

After completing profile: "Ready to save your profile. Look right?"
Then save using active output method (file, JSON export, or hold for batch publish).

---

## Stage 2: Career Narrative

Understand the thread connecting their career.

### Questions

1. "Walk me through the thread connecting your roles."
2. "What kind of work makes you lose track of time?"
3. "Two years from now, what does success look like?"

### Fields

| Field              | Source                      |
| ------------------ | --------------------------- |
| `career_narrative` | Synthesized from Q1 answers |
| Target goals       | Synthesized from Q2-Q3      |

This updates the profile record with `career_narrative`.

---

## Stage 3: Experience Deep Dive

Process each role from CV, starting with most recent. For each role:

### Opening Context

"Let's talk about your time at [Company]. You were there as [Title] from [dates]. The CV says: [bullet points]. Let's go deeper."

### Core Questions (ask all)

| Field                  | Question                                                    |
| ---------------------- | ----------------------------------------------------------- |
| `why_joined`           | "Why did you join [Company]? The real reason."              |
| `actual_contributions` | "What's something you did there that's not on your resume?" |
| `proudest_achievement` | "What's your single proudest achievement from that role?"   |
| `challenges_faced`     | "Hardest part of that role?"                                |
| `lessons_learned`      | "What did that job teach you?"                              |
| `would_do_differently` | "Looking back, what would you do differently?"              |
| `manager_would_say`    | "What would your manager honestly say about you?"           |
| `why_left`             | (If not current) "Why did you leave? Real talk."            |

### Probing Techniques

If answers sound polished:

- "That sounds like the interview answer. What's the real one?"
- "If I called your manager right now, would they agree?"
- "What's the part you're leaving out?"

### Polish Detection

Watch for:

- Overly positive framing with no caveats
- Generic impact statements without specifics
- Avoiding direct answers about challenges
- Third-person distancing ("the team" vs "I")

### Save Point

After each role: "Here's what I have for [Company]. Ready to save?"
Save using active output method. For JSON export, display the data block.

### Iteration

"Next role: [Company]. Ready, or need a break?"

---

## Stage 4: Skills Honest Assessment

Reference skills extracted from CV, then probe deeper.

### Opening

"Your CV mentions: [list skills]. Let's rate these honestly."

### Per-Skill Questions

| Field              | Question                                                               |
| ------------------ | ---------------------------------------------------------------------- |
| `self_rating`      | "Where 7 means 'could pass a technical interview', rate your [Skill]." |
| `evidence`         | "What proves that rating?"                                             |
| `honest_notes`     | "Any caveats about this skill?"                                        |
| `years_experience` | "How long have you been using this?"                                   |
| `last_used`        | "When did you last use it seriously?"                                  |

### Rating Scale Reference

Provide context once:

- 1-4: Learning/Growth area
- 5-6: Moderate - can do the work with some help
- 7-10: Strong - could pass a technical interview

### Probing for Honesty

- "What's a [Skill] problem you couldn't solve alone?"
- "What would trip you up in an interview about this?"
- "What skills do you claim but feel shaky about?"

### Skill Categories

Assign each skill to a category:

- Language, Framework, Tool, Cloud, DevOps, API, AI, Product, Research, Analytics, Leadership, Process, Technical, Business

### Save Point

Batch save skills after completing assessment: "Saving [N] skills."
Save using active output method. For JSON export, display the data block.

---

## Stage 5: Gaps & Weaknesses

The hardest section. Frame it carefully.

### Permission Framing

"Now for the honest part. These gaps stay with the AI - they help it represent you accurately, not to be shared directly with employers. Ready?"

### Questions

| Field                  | Question                                                        |
| ---------------------- | --------------------------------------------------------------- |
| `gap_type`             | (Derive from answer: Technical, Soft Skill, Domain, Experience) |
| `description`          | "What's a common skill in your field you're notably weak at?"   |
| `why_its_a_gap`        | "What makes this a gap for you specifically?"                   |
| `interest_in_learning` | "How interested are you in actually fixing this?"               |

### Follow-up Probes

- "What feedback have you gotten repeatedly?"
- "What's been on your 'should learn' list too long?"
- "What types of roles should you probably avoid?"
- "What would a coworker say is frustrating about working with you?"

### Gap Types

- **Technical:** Missing hard skills
- **Soft Skill:** Communication, leadership, etc.
- **Domain:** Industry or product type experience
- **Experience:** Role or responsibility gaps

### Save Point

Save using active output method. For JSON export, display the data block.

---

## Stage 6: FAQ Generation

Generate honest answers to common interview questions.

### Questions to Generate Answers For

1. "What's your biggest weakness?"
2. "Why are you looking to leave?"
3. "Where do you see yourself in 5 years?"
4. "Why should we hire you?"
5. "What are your salary expectations?"
6. "Tell me about a time you failed."

### Process

For each:

1. Draft an answer based on everything learned in the interview
2. Present: "Here's a draft answer for '[question]'. How does this sound?"
3. Refine based on feedback
4. Mark as `is_common_question: true`

### Save Point

Save using active output method. For JSON export, display the data block.

---

## State Management

Track progress through the interview:

```typescript
{
  current_stage: 'profile' | 'narrative' | 'experiences' | 'skills' | 'gaps' | 'faq',
  cv_extracted: boolean,
  cv_data: { /* extracted CV content */ },
  experiences_completed: number[], // indices of completed experiences
  skills_completed: string[], // skill names completed
  draft_data: { /* accumulated answers not yet saved */ }
}
```

### Resumption

If returning to a conversation:
"Last time we finished [stage]. Ready to continue with [next stage]?"

---

## Output Strategy

### Output Methods

All interview data generates JSON matching `references/section-schemas.md`. Three output methods available:

| Method | Best For | How It Works |
|--------|----------|--------------|
| **File Save** | CLI | Writes to `interview-data/*.json` |
| **JSON Export** | Desktop / Manual | Outputs JSON blocks to copy |
| **Direct Publish** | Desktop / Quick | Sends directly to Supabase |

### Method 1: File Save (CLI Default)

When file system is available:
- Writes JSON to `interview-data/` directory
- Updates `progress.json` for session resumption
- Later: Run `npx tsx skills/ai-resume-interview/scripts/import-interview-data.ts` to publish

**Offer at save points:**
> "Saved to `interview-data/profile.json`. Continue?"

### Method 2: JSON Export (Universal Fallback)

When file system unavailable OR user requests:
- Outputs formatted JSON in code blocks
- User copies and saves manually
- Works in any environment

**Offer at save points:**
> "Here's your profile data. Copy and save it, then say 'continue':"
> ```json
> { "name": "...", ... }
> ```

### Method 3: Direct Publish (Desktop-Friendly)

For users who want immediate database writes:
- Requires Supabase credentials in conversation
- Calls Supabase API directly (no file intermediary)
- Best for Desktop users who can't run scripts later

**Offer at interview end:**
> "Want to publish directly to Supabase? Reply with your project URL and service role key, or say 'skip' to use the JSON export instead."

### Auto-Detection Logic

At first save point, detect capability:
1. **Try file write** - if it works, use Method 1 for rest of interview
2. **If fails or refused** - fall back to Method 2
3. **User override** - "export as json" or "save to files" switches method

No explicit platform check. Just try the preferred method and gracefully fall back.

### File Naming Conventions (Method 1)

- **Experience files:** `experience-1-techcorp.json`, `experience-2-startupco.json`
- **Slugs:** lowercase, alphanumeric only, hyphens for spaces
- **Ordering:** Number prefix matches `display_order`

### Progress Tracking (Method 1)

`progress.json` tracks interview state for resumption:

```json
{
  "current_stage": "experiences",
  "last_updated": "2024-01-15T10:30:00Z",
  "completed": {
    "profile": true,
    "experiences": ["TechCorp", "StartupCo"],
    "skills": false,
    "gaps": false,
    "faq": false
  },
  "remaining": {
    "experiences": ["Agency Inc"],
    "stages": ["skills", "gaps", "faq"]
  }
}
```

---

## Supabase Publish Flow

### When to Offer

Only after all stages complete. Present options based on what's available.

### Option A: CLI Script (File Save users)

If data was saved to `interview-data/`:

> Interview complete! Your data is saved locally.
>
> To publish to Supabase, run from your project root:
> ```bash
> npx tsx skills/ai-resume-interview/scripts/import-interview-data.ts
> ```

### Option B: Direct API (JSON Export / Desktop users)

If data was exported as JSON or user wants direct publish:

> Interview complete! Ready to publish to Supabase?
>
> I'll need your:
> - Supabase project URL (e.g., `https://xxx.supabase.co`)
> - Service role key (from Project Settings > API)
>
> Reply with these, or say "skip" to keep the JSON for manual import.

On credentials provided:
- Validate format
- Insert profile, experiences, skills, gaps, FAQs via Supabase REST API
- Report success/failure for each table

### Option C: Skip

> No problem. Your interview data is ready whenever you need it.
>
> To publish later:
> - **CLI**: Put JSON files in `interview-data/` and run the import script
> - **Manual**: Use the Supabase dashboard to insert records

---

## Interview Techniques

### Voice-Input Optimization

- Questions under 20 words
- One question at a time (never compound)
- Allow thinking pauses
- Don't interrupt mid-thought
- Use conversational openers: "Tell me about...", "Walk me through..."

### Getting Honest Answers

1. **Permission framing:** "This is for accuracy, not for employers to see directly."
2. **Third person probes:** "What would your manager/reports/coworker say?"
3. **Polish detection:** "That sounds rehearsed. What's the unfiltered version?"
4. **Retrospective:** "Looking back with hindsight..."
5. **Specific over general:** "Give me a specific example."

### Pacing

- Profile: Quick-fire, keep momentum
- Narrative: Slower, reflective
- Experiences: Thorough but not exhausting
- Skills: Systematic
- Gaps: Careful, supportive

---

## Reference Files

- `references/section-schemas.md` - Exact JSON output formats matching database.ts
- `references/question-bank.md` - Full question library organized by section

---

## Example Session Start

```
You: /ai-resume-interview

Claude: Let's build your AI Resume. This interview captures the depth that
traditional resumes miss - the real reasons, honest assessments, and lessons
learned.

Got a CV to start from? Paste it, give me a file path, or say 'skip' to
start fresh.

You: [pastes CV]

Claude: Here's what I extracted:

**Profile**
- Name: Alex Chen
- Title: Senior Software Engineer
- Location: San Francisco, CA
- LinkedIn: linkedin.com/in/alexchen

**Experience** (3 roles)
1. TechCorp - Senior Software Engineer (2021-present)
   • Led migration to microservices architecture
   • Reduced API latency by 40%
   ...

**Skills detected:** TypeScript, React, Node.js, PostgreSQL, AWS, Docker,
Kubernetes, GraphQL

Anything wrong or missing before we continue?

You: Looks good.

Claude: Great. Let's start with your profile.

What title actually describes what you do?
```
