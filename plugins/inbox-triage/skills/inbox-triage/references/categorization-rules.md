# Email Categorization Rules

Detailed heuristics for categorizing emails in inbox triage.

## Category Definitions

### 🔴 Urgent

**Definition:** Requires immediate attention (within hours, not days).

**Detection signals:**

| Signal | Examples | Weight |
|--------|----------|--------|
| Explicit urgency | "urgent", "ASAP", "emergency", "critical" | High |
| Near deadline | "by EOD", "today", "in the next hour" | High |
| Priority sender | CEO, direct manager, key client | High |
| Escalation language | "escalating", "blocking", "waiting on you" | High |
| Reply-needed + time | "please confirm by...", "need response before..." | High |

**Priority sender list (customize per user):**
- C-level executives at own company
- Direct manager
- Key clients (by domain or email)
- Legal/compliance/HR for sensitive matters

---

### 🟡 Needs Response

**Definition:** Requires a reply from you, but not immediately.

**Detection signals:**

| Signal | Examples | Weight |
|--------|----------|--------|
| Direct question | "Can you...", "What do you think...", "Do you have..." | Medium |
| Action request | "Please review", "I need your input", "Could you send..." | Medium |
| Awaiting input | "Let me know when...", "Once you've...", "When you have time..." | Medium |
| To: field (not CC) | You're the primary recipient | Medium |
| Thread continuation | Reply expected in ongoing conversation | Medium |

**Exclude if:**
- You're only CC'd
- It's a broadcast/announcement
- No question or action requested

---

### 📋 Admin

**Definition:** Routine business tasks, non-urgent but needs handling.

**Detection signals:**

| Signal | Examples | Weight |
|--------|----------|--------|
| Scheduling | "Can we find a time...", calendar invites | Medium |
| Approvals | "Please approve", expense reports, time-off requests | Medium |
| Documentation | "Please sign", "Review and confirm", "For your records" | Medium |
| Routine requests | "Monthly report", "Regular update", "Standing request" | Low |

---

### 📢 FYI

**Definition:** Informational only, no action required from you.

**Detection signals:**

| Signal | Examples | Weight |
|--------|----------|--------|
| CC only | You're in CC, not To | High |
| Announcement format | "Team announcement", "FYI:", "Update:" | High |
| No question/action | Purely informational content | Medium |
| Broadcast indicators | Multiple recipients, company-wide | Medium |
| Status updates | "Project update", "Weekly summary" | Medium |

---

### 📰 Newsletter

**Definition:** Subscribed content, marketing, digests.

**Detection signals:**

| Signal | Examples | Weight |
|--------|----------|--------|
| Unsubscribe link | Footer contains "unsubscribe" | High |
| Marketing sender | "marketing@", "news@", "updates@" | High |
| Templated design | HTML heavy, promotional layout | Medium |
| Digest format | "Weekly digest", "Top stories", "Newsletter" | Medium |
| Subscription services | Substack, Mailchimp, ConvertKit patterns | Medium |

**Common newsletter domains:**
- substack.com
- mailchimp.com
- *@marketing.*
- *@newsletter.*
- digest@*

---

### 🗑️ Noise

**Definition:** Automated notifications, spam, or irrelevant.

**Detection signals:**

| Signal | Examples | Weight |
|--------|----------|--------|
| Automated sender | "noreply@", "notifications@", "automated@" | High |
| System notifications | GitHub, Jira, Slack email digests | High |
| Receipts/confirmations | "Your order", "Receipt for", "Confirmation" | Medium |
| Social notifications | LinkedIn, Twitter, Facebook emails | Medium |
| Promotional | Sales pitches from unknown senders | Medium |

**Common noise patterns:**
- `noreply@*` - Automated system emails
- `notifications@github.com` - Unless you're actively coding
- `*@jira.*` - Project management noise
- `digest@slack.com` - Slack digests (already seen in Slack)

---

## Decision Tree

Apply in this order:

```
1. Is sender in my priority contacts?
   → Yes: URGENT

2. Contains urgency keywords AND deadline < 24h?
   → Yes: URGENT

3. Is it from noreply/automated sender?
   → Yes: Check if actionable → NOISE or ADMIN

4. Contains unsubscribe link OR newsletter patterns?
   → Yes: NEWSLETTER

5. Am I only CC'd (not in To)?
   → Yes: FYI

6. Does it ask a direct question or request action from me?
   → Yes: NEEDS RESPONSE

7. Is it routine business (scheduling, approvals)?
   → Yes: ADMIN

8. Default
   → FYI
```

---

## Keyword Lists

### Urgency Keywords (case-insensitive)

**High urgency:**
```
urgent, asap, emergency, critical, immediately, right away,
blocking, blocker, escalating, escalation, time-sensitive
```

**Deadline indicators:**
```
by eod, by cob, end of day, today, this morning, this afternoon,
within the hour, in the next, before [time], deadline
```

**Escalation language:**
```
waiting on you, blocked by, need your response, please confirm,
haven't heard back, following up again, second request
```

### Skip Keywords

**Auto-archive worthy:**
```
unsubscribe, manage preferences, email preferences, opt out,
you received this because, this is an automated message
```

---

## Customization Notes

Users should customize:

1. **Priority senders:** Add key clients, executives, important contacts
2. **Skip senders:** Newsletters to always archive, noisy systems
3. **Project keywords:** Terms that always elevate priority
4. **Time zones:** Adjust "EOD" interpretation

Store customizations in user preferences or provide during triage setup.
