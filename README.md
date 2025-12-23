# Claude Skills

A collection of skills for Claude Code to extend its capabilities with specialized functionality.

## What are Claude Skills?

Skills are packages that provide Claude Code with domain-specific knowledge, workflows, and tools. They can include:
- Python scripts with isolated virtual environments
- Reference documentation and guides
- Structured workflows for complex tasks
- Tool integrations and automations

## Available Skills

### PowerPoint Template Updater (`pptx-skills`)

Update PowerPoint templates with new content while preserving formatting, structure, and layout. Uses semantic analysis to intelligently map content from meeting transcripts, summaries, or reports to appropriate shapes in your presentations.

**Use cases:**
- Weekly status updates from meeting notes
- Quarterly business reviews from data exports
- Sales deck customization
- Event presentations from event details

**Features:**
- Semantic template analysis
- Intelligent content mapping
- Formatting preservation
- Bullet hierarchy maintenance
- Length constraint validation
- Isolated virtual environment for secure Python execution

## Installation

### Quick Install (Recommended)

Install skills directly from this marketplace using Claude Code:

```bash
# 1. Add the Fenrir Labs skills marketplace
/plugin marketplace add fenrirlabsnl/claude-skills

# 2. Install the PowerPoint skills plugin
/plugin install pptx-skills@fenrirlabs-skills
```

That's it! The skill will be automatically available to Claude Code.

### Manual Installation

Alternatively, clone the repository and install dependencies manually:

```bash
# 1. Clone this repository
git clone https://github.com/fenrirlabsnl/claude-skills.git
cd claude-skills/skills/pptx-template-updater

# 2. Run the setup script to create isolated environment
./setup.sh

# 3. Copy to your personal skills directory
mkdir -p ~/.claude/skills
cp -r . ~/.claude/skills/pptx-template-updater
```

## Using Skills with Claude Code

Skills are **automatically invoked** by Claude based on your requests - you don't need to call them explicitly.

**Example usage:**
```
You: "Can you update this quarterly report template with data from our meeting notes?"

Claude: [Automatically uses pptx-template-updater skill to analyze and update the presentation]
```

For detailed usage instructions, see the individual skill's SKILL.md file in the `skills/` directory.

## Marketplace Structure

This repository is configured as a Claude Code plugin marketplace:

```
.claude-plugin/
  └── marketplace.json    # Marketplace configuration
skills/
  └── pptx-template-updater/
      ├── SKILL.md        # Skill instructions for Claude
      ├── setup.sh        # Virtual environment setup
      ├── requirements.txt
      ├── scripts/
      └── references/
```

## Contributing

Contributions are welcome! If you've created a skill that could benefit others:

1. Fork this repository
2. Add your skill to the `skills/` directory
3. Include a SKILL.md with clear documentation (see structure below)
4. Add your skill to `.claude-plugin/marketplace.json`
5. Submit a pull request

### Skill Structure

Each skill should follow this structure:

```
skills/
└── your-skill-name/
    ├── SKILL.md              # Main documentation (required)
    ├── setup.sh              # Setup script (if needed)
    ├── requirements.txt      # Dependencies (if needed)
    ├── .gitignore           # Ignore build artifacts
    ├── scripts/             # Executable scripts
    │   └── *.py, *.sh
    └── references/          # Reference documentation
        └── *.md
```

### SKILL.md Requirements

Each SKILL.md must include YAML frontmatter:

```yaml
---
name: your-skill-name
description: What the skill does and when Claude should use it (include trigger keywords)
---

# Skill Name

## Instructions for Claude
Clear step-by-step guidance...
```

**Key guidelines:**
- `name`: lowercase letters, numbers, hyphens only (max 64 chars)
- `description`: Be specific about WHAT it does and WHEN to use it (max 1024 chars)
- Include keywords users might mention to trigger the skill

## Security

Skills use isolated virtual environments to ensure:
- No system-wide package pollution
- Controlled dependency versions
- Easy cleanup and removal
- Reproducible execution environments

## License

Each skill may have its own license. See individual skill directories for details.

## Resources

- [Claude Code Skills Documentation](https://code.claude.com/docs/en/skills.md)
- [Plugin Development Guide](https://code.claude.com/docs/en/plugins.md)
- [Marketplace Configuration](https://code.claude.com/docs/en/plugin-marketplaces.md)
