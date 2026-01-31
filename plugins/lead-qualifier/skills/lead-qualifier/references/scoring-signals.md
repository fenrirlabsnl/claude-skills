# Lead Scoring Signal Detection

Detailed heuristics for detecting and scoring lead qualification signals.

## Budget Signals (+3 points)

**Definition:** Indicators that the prospect has allocated or is discussing budget.

### Explicit Budget Mentions

| Pattern | Example | Confidence |
|---------|---------|------------|
| Dollar amounts | "$10k", "$50,000/year", "100k budget" | High |
| Budget keyword | "budget allocated", "have budget", "budget approved" | High |
| Pricing inquiry | "what's the pricing", "cost for", "how much" | Medium |
| Investment language | "willing to invest", "looking to spend" | Medium |

### Implicit Budget Signals

| Pattern | Example | Confidence |
|---------|---------|------------|
| Procurement mention | "procurement process", "vendor approval" | High |
| Contract language | "contract terms", "agreement structure" | High |
| Comparison shopping | "comparing pricing", "getting quotes" | Medium |
| ROI questions | "what's the ROI", "payback period" | Medium |

### Keywords to Detect

```
budget, pricing, cost, price, investment, spend, afford,
procurement, purchasing, quote, proposal, contract,
roi, return on investment, payback, total cost
```

---

## Timeline Signals (+2 points)

**Definition:** Indicators of a specific implementation timeframe.

### Explicit Timeline

| Pattern | Example | Confidence |
|---------|---------|------------|
| Specific date | "by March 15", "end of Q1" | High |
| Quarter reference | "Q1 implementation", "this quarter" | High |
| ASAP language | "need ASAP", "urgent need", "immediately" | High |
| Month reference | "next month", "by February" | Medium |

### Implicit Timeline

| Pattern | Example | Confidence |
|---------|---------|------------|
| Event-driven | "before our launch", "ahead of expansion" | High |
| External deadline | "audit coming up", "contract renewal" | High |
| Planning language | "planning for Q2", "roadmap for next year" | Medium |
| Seasonal | "before holiday season", "fiscal year end" | Medium |

### Keywords to Detect

```
asap, urgent, immediately, timeline, deadline, by [date],
q1, q2, q3, q4, this quarter, next quarter, this month,
before, ahead of, planning for, need by
```

### Negative Timeline Signals (-2 points)

```
no timeline, just exploring, sometime next year,
no rush, when we're ready, eventually, down the road
```

---

## Decision-Maker Signals (+2 points)

**Definition:** Indicators that the contact has authority to purchase.

### Title-Based Detection

**High confidence (C-level, Founders):**
```
ceo, cto, cfo, coo, cmo, cro, cio, ciso,
chief [anything] officer,
founder, co-founder, cofounder,
owner, partner, principal
```

**Medium confidence (VP/Director):**
```
vp, vice president, v.p.,
svp, senior vice president,
evp, executive vice president,
director, head of, lead
```

**Context-dependent (Manager):**
```
manager, senior manager
→ Score +1 only if paired with decision language
```

### Language-Based Detection

| Pattern | Example | Confidence |
|---------|---------|------------|
| Decision authority | "I'm the decision maker", "I decide" | High |
| Budget authority | "I can approve", "my budget", "I control" | High |
| Team leadership | "my team", "I manage", "I oversee" | Medium |
| Evaluation lead | "I'm evaluating", "leading the selection" | Medium |

### Negative Signals

| Pattern | Example | Confidence |
|---------|---------|------------|
| Delegation | "on behalf of", "my boss asked", "passing along" | High |
| Junior titles | "intern", "associate", "coordinator", "assistant" | High |
| Student | "student", "@edu", "@student" | High |

---

## Company Size Signals (+2 points for 10-500)

**Definition:** Indicators of company size (sweet spot: 10-500 employees).

### Direct Indicators

| Pattern | Confidence |
|---------|------------|
| Explicit count ("team of 50") | High |
| LinkedIn company page employee count | High |
| "We have X people" | High |

### Indirect Indicators

| Pattern | Estimated Size |
|---------|---------------|
| "Just me", "solo", "independent" | 1 |
| "Small team", "startup" | 2-10 |
| "Growing team", "expanding" | 10-50 |
| "Midsized", "Series A/B" | 50-200 |
| "Enterprise", "Fortune 500" | 500+ |
| Multiple office locations mentioned | 100+ |
| "Global team" | 200+ |

### Domain-Based Signals

| Domain Pattern | Signal |
|----------------|--------|
| Personal email (@gmail, @yahoo) | Unknown/small |
| Startup domains (shorter, trendy) | Likely <100 |
| Enterprise domains (bank, insurance) | Likely 500+ |

### Scoring by Size

| Size | Points |
|------|--------|
| 1 (solo) | +0 |
| 2-9 | +1 |
| 10-500 | +2 |
| 500+ | +1 |

---

## Specific Use Case Signals (+1 point)

**Definition:** Prospect describes their actual problem or requirement.

### Strong Use Case Indicators

| Pattern | Example |
|---------|---------|
| Problem description | "We're struggling with...", "Our challenge is..." |
| Feature mention | "We need [specific feature]", "Looking for [capability]" |
| Integration need | "Needs to work with [system]", "Integration with..." |
| Volume/scale | "We process X per month", "Supporting Y users" |

### Weak Use Case Indicators

| Pattern | Example |
|---------|---------|
| Generic interest | "Interested in learning more" |
| Category only | "Looking for a CRM" (no specific needs) |
| No context | "Tell me about your product" |

---

## Bonus Signals

### Referral Mention (+2 points)

```
[name] referred me, recommended by, heard about you from,
[name] said I should reach out, suggested I contact
```

### Booked a Call (+2 points)

- Calendly confirmation
- Meeting request in email
- "Scheduled time" language

### Competition Mention (+1 point)

```
currently using [competitor], comparing to [competitor],
looking at alternatives to, switching from
```

### Reply to Outreach (+1 point)

- "Re:" in subject
- References your previous email
- Thread continuation

---

## Negative Signals

### Generic Exploration (-2 points)

```
just curious, exploring options, no timeline,
just learning, research phase, not sure yet,
early stages, preliminary
```

### Personal/Student (-1 point)

- @gmail.com, @yahoo.com, @hotmail.com
- @*.edu, @student.*
- No company signature

### Unsubscribe Request (-5 points)

```
unsubscribe, remove me, stop emailing,
take me off, don't contact
```

---

## Scoring Summary Table

| Signal | Points | Key Detection |
|--------|--------|---------------|
| Budget/pricing mention | +3 | $, budget, pricing, cost |
| Clear timeline | +2 | Date, quarter, ASAP |
| Decision-maker title | +2 | C-level, VP, Director, Head of |
| Company 10-500 | +2 | Employee count, team size |
| Specific use case | +1 | Problem description, feature need |
| Booked call | +2 | Calendly, meeting confirmation |
| Referral | +2 | "[Name] referred me" |
| Reply to outreach | +1 | "Re:", thread reference |
| Competition mention | +1 | "[Competitor]" comparison |
| Just exploring | -2 | "no timeline", "curious" |
| Personal email | -1 | @gmail, @yahoo |
| Unsubscribe | -5 | "remove me", "unsubscribe" |
