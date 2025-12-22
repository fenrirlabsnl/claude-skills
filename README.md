# Claude Skills

A collection of skills for Claude Code to extend its capabilities with specialized functionality.

## What are Claude Skills?

Skills are packages that provide Claude Code with domain-specific knowledge, workflows, and tools. They can include:
- Python scripts with isolated virtual environments
- Reference documentation and guides
- Structured workflows for complex tasks
- Tool integrations and automations

## Available Skills

### [pptx-template-updater](./pptx-template-updater)

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

## Installation

Each skill has its own installation instructions. Generally:

1. Clone this repository:
```bash
git clone https://github.com/[your-username]/claude-skills.git
cd claude-skills
```

2. Navigate to the skill directory:
```bash
cd pptx-template-updater
```

3. Run the setup script:
```bash
./setup.sh
```

## Using Skills with Claude Code

Skills can be used directly in Claude Code conversations. Reference the skill's SKILL.md file for specific usage instructions and workflows.

## Contributing

Contributions are welcome! If you've created a skill that could benefit others:

1. Fork this repository
2. Add your skill in a new directory
3. Include a SKILL.md with clear documentation
4. Add installation/setup scripts if needed
5. Submit a pull request

## Skill Structure

Each skill should follow this structure:

```
skill-name/
├── SKILL.md              # Main documentation
├── setup.sh              # Setup script (if needed)
├── requirements.txt      # Dependencies (if needed)
├── .gitignore           # Ignore build artifacts
├── scripts/             # Executable scripts
│   └── *.py, *.sh
└── references/          # Reference documentation
    └── *.md
```

## Security

Skills use isolated virtual environments to ensure:
- No system-wide package pollution
- Controlled dependency versions
- Easy cleanup and removal
- Reproducible execution environments

## License

Each skill may have its own license. See individual skill directories for details.
