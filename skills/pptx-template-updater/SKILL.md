---
name: pptx-template-updater
description: Update PowerPoint templates with new content while preserving formatting, structure, and layout. Uses semantic analysis to understand template structure and intelligently map new content (from meeting transcripts, summaries, reports) to appropriate shapes. Use when the user asks to update a PowerPoint presentation, fill in a PPTX template, update slides with new data, or refresh a presentation with new content. Handles business reports, sales decks, status updates, and other template-based presentations.
---

# PowerPoint Template Updater

Update PowerPoint templates by intelligently mapping new content to existing shapes while preserving all formatting, bullet structures, and layout.

## Overview

This skill enables semantic template updates: instead of rigid placeholders, it analyzes the template structure to understand what each shape contains, then intelligently updates text with new content from sources like meeting transcripts, summaries, or structured data.

**Key capabilities:**
- Extract and analyze template structure (shape positions, content, purpose)
- Derive semantic meaning from shape content and positioning
- Preserve bullet point hierarchy and formatting
- Respect text length constraints to prevent overflow
- Update only text content—never modify layouts or formatting

## Installation

**Prerequisites:**
- Python 3.8 or higher
- pip (Python package installer)

**Setup:**

1. Navigate to the skill directory:
```bash
cd pptx-template-updater
```

2. Run the setup script to create an isolated virtual environment and install dependencies:
```bash
./setup.sh
```

This will:
- Create a `.venv` directory with an isolated Python environment
- Install python-pptx and all required dependencies
- Set up the skill for secure, isolated execution

**Note:** The setup only needs to be run once. All scripts will automatically use the isolated environment.

## Workflow

### 1. Analyze Template Structure

First, understand the template you're working with:

```bash
./scripts/run_extract.sh template.pptx --output structure.json
```

This produces a JSON file with:
- All slides and shapes
- Text content and character counts
- Shape positions and types
- Paragraph and bullet information

**Review the structure** to understand:
- What content exists in each shape
- Which shapes should be updated
- Text length constraints

### 2. Understand Semantic Meaning

Use heuristics to infer what each shape represents:

**Consult [semantic-analysis.md](references/semantic-analysis.md)** for detailed guidance on:
- Inferring slide types (title slide, metrics slide, content slide)
- Determining shape purpose from position and content
- Detecting patterns (label-value pairs, temporal content, hierarchical bullets)
- Matching new content to appropriate shapes

**Key inference signals:**
- **Position**: Top shapes = titles, bottom = footers, center = main content
- **Content patterns**: Dates, metrics (numbers/percentages), headers, bullet lists
- **Context**: Surrounding shapes, slide type, existing text patterns

### 3. Extract and Map Content

Parse the new content source (transcript, summary, data file):

```python
# Example: Extract structured data from a meeting transcript
new_data = {
    "date": "Q4 2024",
    "metrics": [
        {"label": "Revenue", "value": "$2.5M"},
        {"label": "Users", "value": "10,000"}
    ],
    "key_points": [
        "Launched new feature X",
        "Expanded to 3 new markets",
        "Customer satisfaction up 15%"
    ]
}
```

Match extracted data to template shapes based on semantic analysis:
- Date fields → update with current date/period
- Metric values → update numbers while preserving format
- Bullet lists → update content while preserving hierarchy
- Headers → update section titles

### 4. Validate Length Constraints

**Critical**: Text must fit within existing shapes.

For each update:
1. Check original text length (from structure analysis)
2. Compare with new text length
3. If new text is >50% longer, either:
   - Summarize/truncate to fit
   - Warn the user
   - Ask for guidance

**See [semantic-analysis.md](references/semantic-analysis.md#length-management)** for text fitting strategies.

### 5. Create Update Instructions

Build a JSON file with update instructions:

```json
{
  "updates": [
    {
      "slide": 1,
      "shape": 2,
      "text": "Q4 2024 Results",
      "preserve_bullets": false
    },
    {
      "slide": 2,
      "shape": 5,
      "text": "Launched new feature X\nExpanded to 3 new markets\nCustomer satisfaction up 15%",
      "preserve_bullets": true
    }
  ]
}
```

### 6. Apply Updates

```bash
./scripts/run_update.sh template.pptx updates.json output.pptx
```

The script will:
- Apply all updates
- Preserve formatting and bullet structure
- Warn if text length exceeds constraints
- Report errors for invalid slide/shape indices

**Review the output** presentation to ensure:
- Text fits within shapes
- Bullet hierarchy is preserved
- Formatting looks correct
- No content was cut off

## Quick Start Example

**User request:** "Update this quarterly report with data from our team meeting notes"

```bash
# 1. Extract template structure
./scripts/run_extract.sh quarterly_report.pptx -o structure.json

# 2. Analyze the structure.json to understand slide types and shape purposes
# Identify: Slide 1 = title slide, Slide 2 = metrics, Slide 3 = key achievements

# 3. Parse meeting notes to extract:
#    - Current quarter/date
#    - Key metrics (revenue, users, growth)
#    - Achievement bullet points

# 4. Create updates.json mapping content to shapes:
{
  "updates": [
    {"slide": 1, "shape": 2, "text": "Q4 2024 Quarterly Report"},
    {"slide": 2, "shape": 3, "text": "$2.5M"},
    {"slide": 2, "shape": 4, "text": "10,000"},
    {"slide": 3, "shape": 2, "text": "Launched feature X\nExpanded to 3 markets\n15% satisfaction increase", "preserve_bullets": true}
  ]
}

# 5. Apply updates
./scripts/run_update.sh quarterly_report.pptx updates.json updated_report.pptx
```

## Reference Documentation

**For detailed technical guidance:**

- **[python-pptx-guide.md](references/python-pptx-guide.md)**: Complete reference for working with python-pptx library
  - Reading and writing presentations
  - Working with text, paragraphs, and runs
  - Preserving formatting
  - Common pitfalls and best practices

- **[semantic-analysis.md](references/semantic-analysis.md)**: Comprehensive guide to inferring meaning from templates
  - Slide type detection
  - Shape purpose inference
  - Content matching strategies
  - Pattern recognition (label-value pairs, temporal content, hierarchical bullets)
  - Length management and text fitting

## Best Practices

1. **Always analyze first**: Run `extract_template_structure.py` before updating to understand constraints
2. **Preserve structure**: Match the original template's organization—don't add complexity it doesn't have
3. **Respect lengths**: Never overflow shapes; summarize or ask for guidance if text is too long
4. **Maintain hierarchy**: Preserve bullet point levels and paragraph structure
5. **Test incrementally**: Update a few shapes first, review output, then proceed with remaining updates
6. **Handle ambiguity**: When unsure which shape to update, ask the user or provide options
7. **Validate output**: Always review the generated presentation before considering it final

## Common Scenarios

**Scenario 1: Weekly status update from meeting transcript**
- Extract key discussion points and decisions
- Map to "Updates" and "Action Items" bullet lists
- Update date field with current week

**Scenario 2: Quarterly business review from data export**
- Extract metrics (revenue, growth, users)
- Map to metric value shapes (usually near labels)
- Update period indicator (Q1 2024 → Q2 2024)

**Scenario 3: Sales deck customization**
- Replace company name, product names
- Update customer testimonials
- Modify feature lists based on prospect needs

**Scenario 4: Event presentation from event details**
- Update event name, date, location
- Replace speaker names and bios
- Modify agenda/schedule
