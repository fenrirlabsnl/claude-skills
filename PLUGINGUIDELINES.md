# Plugin Guidelines for Claude Code

Comprehensive guide for creating Claude Code plugins. Derived from the official `plugin-dev` reference plugin and the `finance-plugin-example` reference implementation.

---

## Table of Contents

1. [What Is a Plugin](#1-what-is-a-plugin)
2. [Directory Structure](#2-directory-structure)
3. [Plugin Manifest](#3-plugin-manifest)
4. [Commands](#4-commands)
5. [Skills](#5-skills)
6. [Commands vs Skills](#6-commands-vs-skills)
7. [Agents](#7-agents)
8. [MCP Integration](#8-mcp-integration)
9. [Connector Abstraction](#9-connector-abstraction)
10. [Graceful Degradation](#10-graceful-degradation)
11. [Hooks](#11-hooks)
12. [README Structure](#12-readme-structure)
13. [Versioning](#13-versioning)
14. [Quality Checklist](#14-quality-checklist)
15. [Common Mistakes](#15-common-mistakes)

---

## 1. What Is a Plugin

A plugin is a directory-based extension that adds capabilities to Claude Code. Plugins bundle commands, skills, agents, hooks, and MCP server integrations into a distributable package that Claude Code discovers and loads automatically.

**What plugins provide:**
- **Commands** — User-invocable slash commands (`/deploy`, `/review`)
- **Skills** — Background domain knowledge loaded automatically when relevant
- **Agents** — Specialized sub-agents for complex tasks
- **Hooks** — Event-driven automation (pre/post tool use, session start, stop)
- **MCP Servers** — External service integrations (APIs, databases, tools)

**Installation:**
```bash
# From marketplace
/plugin install plugin-name

# Local development
claude --plugin-dir /path/to/plugin
```

**How Claude Code discovers plugins:**
1. Scans enabled plugins for `.claude-plugin/plugin.json`
2. Discovers components in default and custom paths
3. Parses YAML frontmatter and configurations
4. Registers all components (commands, skills, agents, hooks, MCP servers)
5. Initializes MCP servers and hooks

---

## 2. Directory Structure

### Minimal Plugin

```
my-plugin/
├── .claude-plugin/
│   └── plugin.json          # Required — plugin manifest
└── commands/
    └── my-command.md         # At least one component
```

### Standard Plugin

```
my-plugin/
├── .claude-plugin/
│   └── plugin.json          # Plugin manifest (required)
├── .mcp.json                # MCP server configurations (optional)
├── README.md                # Documentation and navigation hub
├── commands/                # User-invocable slash commands
│   └── review.md
├── skills/                  # Background knowledge (auto-loaded)
│   └── domain-knowledge/
│       ├── SKILL.md
│       └── references/
│           └── detailed-guide.md
├── agents/                  # Specialized sub-agents
│   └── code-reviewer.md
└── hooks/                   # Event-driven automation
    ├── hooks.json
    └── scripts/
        └── validate.sh
```

### Advanced Plugin (Enterprise-Grade)

```
enterprise-devops/
├── .claude-plugin/
│   └── plugin.json
├── .mcp.json
├── README.md
├── CONNECTORS.md            # Tool-agnostic connector documentation
├── LICENSE
├── commands/
│   ├── ci/
│   │   ├── build.md
│   │   └── deploy.md
│   └── monitoring/
│       └── status.md
├── agents/
│   ├── orchestration/
│   │   └── deployment-orchestrator.md
│   └── specialized/
│       └── kubernetes-expert.md
├── skills/
│   └── kubernetes-ops/
│       ├── SKILL.md
│       ├── references/
│       │   ├── deployment-patterns.md
│       │   └── troubleshooting.md
│       ├── examples/
│       │   └── basic-deployment.yaml
│       └── scripts/
│           └── validate-manifest.sh
├── hooks/
│   ├── hooks.json
│   └── scripts/
│       ├── security/
│       │   └── scan-secrets.sh
│       └── quality/
│           └── check-config.sh
├── servers/                  # Bundled MCP server source
│   └── custom-mcp/
│       ├── index.js
│       └── package.json
├── lib/                      # Shared utilities
│   ├── core/
│   └── integrations/
└── config/
    └── environments/
        ├── production.json
        └── staging.json
```

**Key conventions:**
- `.claude-plugin/plugin.json` MUST exist at the plugin root — Claude Code won't recognize the plugin without it
- Default directories (`commands/`, `skills/`, `agents/`, `hooks/`, `.mcp.json`) are scanned automatically
- Custom paths can be specified in `plugin.json` to supplement (not replace) defaults
- The `${CLAUDE_PLUGIN_ROOT}` variable resolves to the plugin's absolute path at runtime

---

## 3. Plugin Manifest

**Location:** `.claude-plugin/plugin.json` (required)

### Schema

```json
{
  "name": "my-plugin",
  "version": "1.0.0",
  "description": "Brief description of plugin functionality (50-200 chars)",
  "author": {
    "name": "Author Name",
    "email": "author@example.com",
    "url": "https://example.com"
  },
  "homepage": "https://docs.example.com/my-plugin",
  "repository": "https://github.com/user/my-plugin",
  "license": "MIT",
  "keywords": ["keyword1", "keyword2"],
  "commands": ["./commands", "./admin-commands"],
  "agents": ["./agents/orchestration", "./agents/specialized"],
  "hooks": "./hooks/hooks.json",
  "mcpServers": "./.mcp.json"
}
```

### Field Reference

| Field | Required | Type | Description |
|-------|----------|------|-------------|
| `name` | **Yes** | String | Unique identifier. Kebab-case, letters/numbers/hyphens. Must match `/^[a-z][a-z0-9]*(-[a-z0-9]+)*$/` |
| `version` | No | String | Semantic versioning (MAJOR.MINOR.PATCH). Default: `"0.1.0"` |
| `description` | No | String | 50-200 characters. Active voice. Focus on what, not how |
| `author` | No | Object/String | `{ name, email?, url? }` or `"Name <email> (url)"` |
| `homepage` | No | String (URL) | Plugin documentation URL |
| `repository` | No | String/Object | Source code repository URL |
| `license` | No | String | SPDX identifier (`"MIT"`, `"Apache-2.0"`, etc.) |
| `keywords` | No | Array | 5-10 searchable tags for discovery |
| `commands` | No | String/Array | Additional command directories (supplements `./commands/`) |
| `agents` | No | String/Array | Additional agent directories (supplements `./agents/`) |
| `hooks` | No | String/Object | Hook config file path or inline configuration |
| `mcpServers` | No | String/Object | MCP config file path or inline configuration |

### Path Rules

All paths in `plugin.json` MUST:
- Be relative (no absolute paths)
- Start with `./`
- Use forward slashes only
- Not use `../` (no parent directory navigation)

```
./commands           ← Correct
./src/commands       ← Correct
commands             ← Wrong (missing ./)
/Users/me/commands   ← Wrong (absolute path)
../shared/commands   ← Wrong (parent traversal)
```

### Minimal vs Complete Examples

**Bare minimum** — only `name` is required:
```json
{
  "name": "hello-world"
}
```

**Recommended** — good metadata for distribution:
```json
{
  "name": "code-review-assistant",
  "version": "1.0.0",
  "description": "Automates code review with style checks and suggestions",
  "author": {
    "name": "Jane Developer",
    "email": "jane@example.com"
  },
  "license": "MIT",
  "keywords": ["code-review", "automation", "quality"]
}
```

---

## 4. Commands

Commands are user-invocable slash commands. When a user types `/command-name`, Claude Code loads the command's Markdown content as instructions for Claude.

**Location:** `commands/command-name.md`

### Critical Principle

**Commands are instructions FOR Claude, not messages TO the user.** The command content becomes Claude's directive when invoked.

```markdown
<!-- Correct: instructions for Claude -->
Review this code for security vulnerabilities including:
- SQL injection
- XSS attacks
Provide specific line numbers and severity ratings.

<!-- Wrong: message to the user -->
This command will review your code for security issues.
You'll receive a report with vulnerability details.
```

### Frontmatter

All frontmatter fields are optional. Commands work without any frontmatter.

```markdown
---
description: Review code for security issues
argument-hint: [file-path] [severity-level]
allowed-tools: Read, Grep, Bash(git:*)
model: sonnet
disable-model-invocation: false
---

Command prompt content here...
```

| Field | Type | Description |
|-------|------|-------------|
| `description` | String | Shown in `/help`. Keep under 60 characters. Start with a verb |
| `argument-hint` | String | Documents expected arguments. Use `[brackets]` for each arg |
| `allowed-tools` | String/Array | Restrict tool access. Use `Bash(git:*)` format for Bash filters |
| `model` | String | `haiku` (fast), `sonnet` (balanced), `opus` (complex) |
| `disable-model-invocation` | Boolean | If `true`, only users can invoke (not Claude programmatically) |

### Dynamic Features

**Positional arguments:** `$1`, `$2`, `$3` etc.
```markdown
Deploy $1 to $2 environment using version $3
```
Usage: `/deploy api staging v1.2.3`

**All arguments:** `$ARGUMENTS`
```markdown
Fix issue #$ARGUMENTS following our coding standards.
```
Usage: `/fix-issue 123`

**File references:** `@filepath`
```markdown
Review @$1 for code quality issues.
```
Usage: `/review src/api/users.ts` — Claude reads the file before processing.

**Inline bash execution:** `` !`command` ``
```markdown
Current changes: !`git diff --name-only`
Review each changed file for code quality.
```

**Plugin file references:** `${CLAUDE_PLUGIN_ROOT}`
```markdown
Run analysis: !`node ${CLAUDE_PLUGIN_ROOT}/scripts/analyze.js $1`
Load config: @${CLAUDE_PLUGIN_ROOT}/config/settings.json
```

### Command Workflow Pattern

Well-structured commands follow a workflow:

```markdown
---
description: Analyze financial data for quarterly report
argument-hint: [quarter] [year]
allowed-tools: Read, Grep, Bash(python:*)
---

## Phase 1: Gather Data

Collect financial data for Q$1 $2:
- Revenue figures from @data/revenue.csv
- Expense reports: !`python ${CLAUDE_PLUGIN_ROOT}/scripts/fetch-expenses.py $1 $2`

## Phase 2: Process

Analyze trends, calculate key metrics (YoY growth, margins, burn rate).

## Phase 3: Generate Output

Create the quarterly report following template:
@${CLAUDE_PLUGIN_ROOT}/templates/quarterly-report.md

## Phase 4: Review Checklist

Before delivering:
- [ ] All figures cross-referenced
- [ ] Calculations verified
- [ ] Formatting consistent
- [ ] Disclaimer included
```

---

## 5. Skills

Skills are background domain knowledge that Claude loads automatically when task context matches the skill's trigger description. Unlike commands (user-invoked), skills are model-invoked — Claude decides when to use them.

**Location:** `skills/skill-name/SKILL.md`

### Frontmatter (Required)

```yaml
---
name: financial-analysis
description: This skill should be used when the user asks to "analyze financial data", "create a financial model", "review quarterly figures", mentions "revenue analysis", or discusses financial reporting and metrics.
version: 1.0.0
---
```

| Field | Required | Description |
|-------|----------|-------------|
| `name` | **Yes** | Skill identifier (kebab-case) |
| `description` | **Yes** | Trigger conditions. MUST use third person and include specific phrases |
| `version` | No | Semantic version |

### Description — The Most Important Field

The `description` determines when Claude activates the skill. It must:
- Use third person: `"This skill should be used when..."`
- Include specific trigger phrases in quotes
- List concrete scenarios and keywords
- Be specific, not vague

```yaml
# Good — specific trigger phrases, third person
description: This skill should be used when the user asks to "create a hook", "add a PreToolUse hook", "validate tool use", "implement prompt-based hooks", or mentions hook events (PreToolUse, PostToolUse, Stop).

# Bad — vague, wrong person, no triggers
description: Provides guidance for working with hooks.
```

### Progressive Disclosure

Skills use a three-level loading system to manage context efficiently:

1. **Metadata** (name + description) — Always in context (~100 words)
2. **SKILL.md body** — Loaded when skill triggers (keep under 2,000 words, max 5,000)
3. **Bundled resources** — Loaded as needed by Claude (unlimited)

```
skills/financial-analysis/
├── SKILL.md              # Core skill (1,500-2,000 words ideal)
├── references/           # Loaded as needed
│   ├── schemas.md        # Database schemas, API docs
│   └── patterns.md       # Detailed patterns and techniques
├── examples/             # Copy-paste code samples
│   └── model-template.py
├── scripts/              # Executable utilities
│   └── validate-data.py
└── assets/               # Files used in output (templates, images)
    └── report-template.xlsx
```

### Writing Style

- **Body:** Use imperative/infinitive form (verb-first), not second person
- **Description:** Third person (`"This skill should be used when..."`)
- Reference supporting files so Claude knows they exist

```markdown
## Core Workflow

Parse the financial data from the provided source.
Calculate key metrics: revenue growth, margins, burn rate.
Cross-reference against historical data in `references/schemas.md`.

## Additional Resources

For detailed patterns, consult:
- **`references/schemas.md`** — Database schemas and data formats
- **`references/patterns.md`** — Analysis patterns and techniques
- **`examples/model-template.py`** — Financial model template
```

---

## 6. Commands vs Skills

| Dimension | Commands | Skills |
|-----------|----------|--------|
| **Invocation** | User types `/command-name` | Claude loads automatically based on context |
| **Purpose** | Execute a specific workflow | Provide background knowledge |
| **Trigger** | Explicit user action | Task context matches description |
| **Structure** | Workflow steps with phases | Domain knowledge with references |
| **Length** | Any length (it's a prompt) | Keep lean (1,500-2,000 words in SKILL.md) |
| **Arguments** | Supports `$1`, `$2`, `$ARGUMENTS` | No arguments |
| **Tools** | Can restrict via `allowed-tools` | Cannot restrict tools |
| **Model** | Can override via `model` | Cannot override model |

### Decision Framework

**Use a command when:**
- The user explicitly initiates a workflow (`/deploy`, `/review-pr`)
- The task has a specific input → output flow
- You need to restrict tools or override the model
- Arguments customize each invocation

**Use a skill when:**
- Claude should automatically know something when relevant
- Domain expertise should be available across all commands and conversations
- The knowledge is reference material, not a workflow
- No user action should be required to activate it

**How they complement each other:**
- A `/deploy` command might rely on a `kubernetes-ops` skill for deployment knowledge
- A `/review` command might leverage a `code-standards` skill for project-specific rules
- Commands define *what to do*; skills provide *what to know while doing it*

---

## 7. Agents

Agents are specialized sub-agents that Claude can spawn via the Task tool for complex tasks.

**Location:** `agents/agent-name.md`

### Frontmatter

```yaml
---
name: code-reviewer
description: Reviews code for bugs, logic errors, security vulnerabilities, and adherence to project conventions
tools: Glob, Grep, LS, Read, NotebookRead, WebFetch, TodoWrite, WebSearch
model: sonnet
color: red
---
```

| Field | Required | Description |
|-------|----------|-------------|
| `name` | **Yes** | Agent identifier |
| `description` | **Yes** | What the agent does. Include `<example>` blocks for reliable triggering |
| `tools` | No | Comma-separated list of allowed tools |
| `model` | No | `haiku`, `sonnet`, `opus`, or `inherit` |
| `color` | No | UI color indicator |

### Agent Body Structure

```markdown
You are an expert [role] specializing in [domain].

## Core Responsibilities
1. First responsibility
2. Second responsibility

## Process
Step-by-step approach the agent follows.

## Output Format
How results should be structured and presented.
```

### Agent vs Command vs Skill

- **Agent:** Spawned by Claude for complex sub-tasks. Has its own tool set and model. Runs autonomously.
- **Command:** User-invoked workflow prompt.
- **Skill:** Background knowledge, never spawned.

---

## 8. MCP Integration

Model Context Protocol (MCP) enables plugins to integrate with external services by providing structured tool access.

**Location:** `.mcp.json` at plugin root (or inline in `plugin.json`)

### Server Types

#### stdio — Local Process
```json
{
  "database-tools": {
    "command": "npx",
    "args": ["-y", "@modelcontextprotocol/server-filesystem", "/allowed/path"],
    "env": {
      "LOG_LEVEL": "debug"
    }
  }
}
```
Best for: Custom servers, local tools, NPM-packaged MCP servers.

#### SSE — Server-Sent Events
```json
{
  "asana": {
    "type": "sse",
    "url": "https://mcp.asana.com/sse"
  }
}
```
Best for: Hosted services with OAuth (Asana, Slack, GitHub). OAuth flows handled automatically.

#### HTTP — REST API
```json
{
  "api-service": {
    "type": "http",
    "url": "https://api.example.com/mcp",
    "headers": {
      "Authorization": "Bearer ${API_TOKEN}"
    }
  }
}
```
Best for: Token-authenticated API backends.

#### WebSocket — Real-time
```json
{
  "realtime": {
    "type": "ws",
    "url": "wss://mcp.example.com/ws"
  }
}
```
Best for: Real-time streaming, persistent connections.

### Environment Variables

Use `${VARIABLE_NAME}` for substitution:
```json
{
  "command": "${CLAUDE_PLUGIN_ROOT}/servers/my-server",
  "env": {
    "API_KEY": "${MY_API_KEY}",
    "DATABASE_URL": "${DB_URL}"
  }
}
```

Always use `${CLAUDE_PLUGIN_ROOT}` for portable paths within the plugin.

### MCP Tool Naming

Tools from plugin MCP servers are automatically namespaced:
```
mcp__plugin_<plugin-name>_<server-name>__<tool-name>
```

Pre-allow specific MCP tools in commands:
```yaml
---
allowed-tools: ["mcp__plugin_my-plugin_db__query_data", "mcp__plugin_my-plugin_db__insert_record"]
---
```

Prefer specific tool names over wildcards (`mcp__plugin_name_server__*`) for security.

### Quick Reference Table

| Type | Transport | Best For | Auth |
|------|-----------|----------|------|
| stdio | Process (stdin/stdout) | Local tools, custom servers | Env vars |
| SSE | HTTP streaming | Hosted services, cloud APIs | OAuth |
| HTTP | REST | API backends | Bearer tokens |
| ws | WebSocket | Real-time, streaming | Tokens |

---

## 9. Connector Abstraction

For plugins that integrate with multiple tools in the same category (e.g., multiple data warehouses, multiple CI systems), use connector abstraction to keep the plugin tool-agnostic.

### The Pattern

Instead of hardcoding tool-specific instructions, use category placeholders that the user maps to their specific tools.

**CONNECTORS.md** — Document supported connector categories and their implementations:

```markdown
# Connectors

This plugin uses tool-agnostic connector categories. Map each category
to your specific tool.

## ~~data warehouse

The data warehouse connector. Supports:
- **Snowflake** — via snowflake MCP server
- **BigQuery** — via bigquery MCP server
- **Databricks** — via databricks MCP server

Configure in .mcp.json with the appropriate server for your stack.

## ~~project tracker

The project tracking connector. Supports:
- **Jira** — via jira MCP server
- **Linear** — via linear MCP server
- **Asana** — via asana SSE endpoint
```

### Why This Matters

- Plugin authors write workflows once, not per tool
- Users configure their stack in `.mcp.json` — commands adapt automatically
- Adding new tool support means adding a new MCP server config, not rewriting commands
- Commands reference the *category* ("query the data warehouse"), not the *tool* ("run a Snowflake query")

### Implementation

In commands, reference the category, not the specific tool:

```markdown
## Step 1: Gather Data

Query the ~~data warehouse for quarterly revenue figures.
Use whatever data warehouse tool is configured in the MCP servers.

## Step 2: Track Progress

Create a task in the ~~project tracker to record the analysis.
```

---

## 10. Graceful Degradation

Every command should work both WITH and WITHOUT MCP connections. Users may not have all external services configured, and commands should still be useful.

### The Pattern

Structure commands with explicit connected/disconnected branches:

```markdown
---
description: Analyze quarterly financial data
argument-hint: [quarter] [year]
---

## Step 1: Gather Financial Data

**If data warehouse MCP is connected:**
Query the data warehouse for Q$1 $2 revenue, expenses, and metrics.

**If data warehouse MCP is NOT connected:**
Ask the user to provide financial data. Suggest these sources:
- Export from their accounting system (CSV or Excel)
- Manual entry of key figures
- Point to a local data file they can share

## Step 2: Generate Analysis

[Analysis steps that work regardless of data source]

## Step 3: Deliver Report

**If project tracker MCP is connected:**
Create a task with the completed analysis attached.

**If project tracker MCP is NOT connected:**
Save the report locally and provide the file path.
Suggest the user copy it to their project tracker manually.
```

### Why This Matters

- Plugins work out of the box, even before users configure MCP servers
- Reduces friction for new users evaluating the plugin
- Manual fallback instructions teach users what the automated path does
- No hard failures — every workflow completes, just with different levels of automation

### Guidelines

1. **Always provide manual alternatives** — tell the user how to accomplish the step without the tool
2. **Check, don't crash** — use conditional language ("if connected... if not connected...")
3. **Degrade incrementally** — each disconnected service reduces automation, not functionality
4. **Document what's better with connections** — motivate users to configure MCP servers

---

## 11. Hooks

Hooks are event-driven automation scripts that execute in response to Claude Code events.

**Location:** `hooks/hooks.json` (or inline in `plugin.json`)

### Hook Events

| Event | When It Fires | Use For |
|-------|---------------|---------|
| `PreToolUse` | Before a tool executes | Validation, security scanning |
| `PostToolUse` | After a tool executes | Logging, status updates |
| `Stop` | When Claude finishes a task | Quality checks, notifications |
| `SessionStart` | When Claude Code starts | Permission validation, setup |

### Configuration

```json
{
  "PreToolUse": [
    {
      "matcher": "Write|Edit",
      "hooks": [
        {
          "type": "command",
          "command": "bash ${CLAUDE_PLUGIN_ROOT}/hooks/scripts/scan-secrets.sh",
          "timeout": 30
        }
      ]
    }
  ],
  "Stop": [
    {
      "matcher": ".*",
      "hooks": [
        {
          "type": "prompt",
          "prompt": "Before completing, verify all code meets our quality standards.",
          "timeout": 20
        }
      ]
    }
  ]
}
```

### Hook Types

- **`command`** — Executes a shell command. Returns JSON with `systemMessage` field.
- **`prompt`** — Sends a prompt to Claude for evaluation.

### Script Output Format

Hook scripts should output JSON:
```bash
echo '{"systemMessage": "Validation passed. Ready to commit."}'
exit 0

# Or on failure:
echo '{"systemMessage": "Found 3 security issues. Please fix before completing."}'
exit 1
```

---

## 12. README Structure

Every plugin should include a `README.md` that serves as the navigation hub.

### Template

```markdown
# Plugin Name

Brief description of plugin functionality.

## Overview

What the plugin does, who it's for, and why it exists.

## Installation

How to install the plugin.

## Commands

| Command | Description |
|---------|-------------|
| `/command-1` | Does X |
| `/command-2` | Does Y |

## Skills

| Skill | Activates When |
|-------|---------------|
| `domain-knowledge` | User discusses topic X |
| `best-practices` | User reviews or writes code |

## Agents

| Agent | Role |
|-------|------|
| `code-reviewer` | Deep code review and security analysis |

## Configuration

### Required Environment Variables

| Variable | Description |
|----------|-------------|
| `API_KEY` | API key for service X |

### MCP Server Setup

How to configure the MCP servers this plugin provides.

## Example Workflows

Practical usage examples showing commands, skills, and agents working together.

## Troubleshooting

Common issues and solutions.

## License

License information.
```

---

## 13. Versioning

### Semantic Versioning

Plugin versions follow semantic versioning (MAJOR.MINOR.PATCH):

- **MAJOR** — Breaking changes (removed commands, changed behavior)
- **MINOR** — New functionality, backward-compatible
- **PATCH** — Bug fixes, backward-compatible

```json
{
  "version": "1.0.0"
}
```

Pre-release versions are supported:
```json
"version": "1.0.0-alpha.1"
"version": "1.0.0-beta.2"
"version": "1.0.0-rc.1"
```

### Versioned Directory Pattern

For plugins distributed outside the marketplace where multiple versions may coexist, wrap the plugin in a version directory:

```
my-plugin/
└── 1.0.0/
    ├── .claude-plugin/
    │   └── plugin.json
    ├── commands/
    └── skills/
```

This allows side-by-side installations of different versions. The marketplace handles versioning automatically, so this pattern is primarily for self-distributed plugins.

---

## 14. Quality Checklist

### Pre-Publish Verification

**Manifest:**
- [ ] `.claude-plugin/plugin.json` exists and is valid JSON
- [ ] `name` is kebab-case, unique, starts with letter
- [ ] `version` follows MAJOR.MINOR.PATCH
- [ ] `description` is 50-200 characters, active voice
- [ ] All paths use `./` prefix and forward slashes

**Commands:**
- [ ] Each command is a `.md` file in `commands/`
- [ ] YAML frontmatter is valid (if present)
- [ ] `description` under 60 characters, starts with verb
- [ ] `argument-hint` matches positional arguments used
- [ ] `allowed-tools` uses most restrictive set that works
- [ ] Commands are instructions FOR Claude (not messages to user)
- [ ] Bash filters use `Bash(git:*)` format, not bare `Bash`
- [ ] Graceful degradation: works with and without MCP connections

**Skills:**
- [ ] Each skill has `SKILL.md` in `skills/skill-name/`
- [ ] Frontmatter has `name` and `description` (both required)
- [ ] Description uses third person: `"This skill should be used when..."`
- [ ] Description includes specific trigger phrases in quotes
- [ ] Body uses imperative/infinitive form (not second person)
- [ ] SKILL.md body under 2,000 words (detail in `references/`)
- [ ] All referenced files actually exist
- [ ] No duplicated information between SKILL.md and references

**Agents:**
- [ ] Each agent is a `.md` file in `agents/`
- [ ] Has `description` with triggering examples
- [ ] `tools` list is appropriate for agent's role
- [ ] `model` matches task complexity

**MCP Servers:**
- [ ] Valid JSON in `.mcp.json`
- [ ] Uses `${CLAUDE_PLUGIN_ROOT}` for portable paths
- [ ] Uses `${VARIABLE_NAME}` for secrets (never hardcoded)
- [ ] HTTPS/WSS only (not HTTP/WS)
- [ ] Required environment variables documented in README

**Hooks:**
- [ ] Valid JSON in `hooks/hooks.json`
- [ ] Scripts are executable (`chmod +x`)
- [ ] Scripts output valid JSON with `systemMessage`
- [ ] Reasonable timeouts set
- [ ] Uses `${CLAUDE_PLUGIN_ROOT}` for script paths

**General:**
- [ ] README.md documents all commands, skills, agents
- [ ] README includes installation instructions
- [ ] README documents required environment variables
- [ ] No hardcoded absolute paths anywhere
- [ ] No committed secrets or credentials
- [ ] Test on clean install (not just dev environment)

---

## 15. Common Mistakes

### 1. Weak Skill Trigger Descriptions

```yaml
# Bad — Claude won't know when to activate
description: Provides guidance for financial analysis.

# Good — specific triggers, third person
description: This skill should be used when the user asks to "analyze financial data", "create a financial model", "review quarterly figures", mentions "revenue analysis", or discusses financial reporting.
```

### 2. Commands Written as Messages to Users

```markdown
# Bad — tells the user what will happen
This command will review your code and provide feedback.

# Good — tells Claude what to do
Review the code in the current repository for:
1. Security vulnerabilities
2. Logic errors
3. Style violations
Provide specific file:line references and severity ratings.
```

### 3. Bloated SKILL.md

```
# Bad — 8,000 words in one file, always loaded
skills/my-skill/
└── SKILL.md  (8,000 words)

# Good — progressive disclosure
skills/my-skill/
├── SKILL.md          (1,800 words — core essentials)
└── references/
    ├── patterns.md   (2,500 words — loaded when needed)
    └── advanced.md   (3,700 words — loaded when needed)
```

### 4. Hardcoded Paths

```json
// Bad — breaks on other machines
{
  "command": "/Users/jane/plugins/my-plugin/servers/api-server"
}

// Good — portable
{
  "command": "${CLAUDE_PLUGIN_ROOT}/servers/api-server"
}
```

### 5. No Graceful Degradation

```markdown
# Bad — fails if MCP not configured
Query the database for user metrics.

# Good — works either way
**If database MCP is connected:**
Query the database for user metrics.

**If database MCP is NOT connected:**
Ask the user to provide metrics data from their database export.
```

### 6. Second Person in Skills

```markdown
# Bad — second person
You should start by reading the configuration file.
You need to validate the input before processing.

# Good — imperative form
Start by reading the configuration file.
Validate the input before processing.
```

### 7. Unrestricted Bash in Commands

```yaml
# Bad — allows any bash command
allowed-tools: Bash

# Good — scoped to specific tools
allowed-tools: Bash(git:*), Bash(npm:*), Read, Grep
```

### 8. Hardcoded Secrets in MCP Config

```json
// Bad — secret in source
{
  "headers": {
    "Authorization": "Bearer sk-live-abc123"
  }
}

// Good — environment variable
{
  "headers": {
    "Authorization": "Bearer ${API_TOKEN}"
  }
}
```

### 9. Missing Resource References in Skills

```markdown
# Bad — Claude doesn't know references exist
## Financial Analysis
[Core content only, no mention of references/]

# Good — Claude knows where to find more
## Financial Analysis
[Core content]

## Additional Resources
- **`references/schemas.md`** — Database schemas and data formats
- **`examples/model-template.py`** — Financial model template
```

### 10. Name Conflicts

Plugin names, command names, and skill names should be unique. Name conflicts cause errors at load time. Use descriptive, specific names rather than generic ones like `test`, `run`, or `build`.
