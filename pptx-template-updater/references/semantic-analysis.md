# Semantic Analysis of PowerPoint Templates

Guide for inferring meaning from shape content, position, and structure to intelligently update templates.

## Core Principle

Instead of rigid placeholders like `{{Company Name}}`, this approach derives semantic meaning from:
1. **Shape content** - What text currently exists
2. **Shape position** - Where on the slide it appears
3. **Shape context** - Surrounding shapes and slide structure
4. **Text patterns** - Headers, lists, dates, metrics

## Analysis Workflow

### 1. Extract Template Structure

First, understand the template's complete structure:

```python
# Use extract_template_structure.py to get JSON output
structure = extract_template_structure("template.pptx")

# Analyze slides to identify types
for slide in structure["slides"]:
    slide_type = infer_slide_type(slide)
    for shape in slide["shapes"]:
        shape_purpose = infer_shape_purpose(shape, slide_type)
```

### 2. Infer Slide Types

Common slide type patterns:

| Content Pattern | Likely Type | Common Shapes |
|----------------|-------------|---------------|
| Single large text box, minimal bullets | Title Slide | Company name, presentation title, date |
| Header + numbered/bulleted list | Content Slide | Section title, key points |
| Header + multiple columns | Comparison Slide | Feature comparison, pros/cons |
| Header + large numbers | Metrics Slide | KPIs, statistics, growth numbers |
| Minimal text, mostly whitespace | Divider Slide | Section header, transition |
| Dense bullet points | Detail Slide | Action items, specifications |

**Heuristics:**
```python
def infer_slide_type(slide_data):
    shapes = slide_data["shapes"]

    # Title slide: Few shapes, large text boxes
    if len(shapes) <= 3:
        return "title_slide"

    # Metrics slide: Contains numbers, percentage signs
    text = " ".join([s["text_content"] for s in shapes])
    if re.search(r'\d+%|\$\d+|^\d+$', text):
        return "metrics_slide"

    # Content slide: Many bullets
    total_bullets = sum(s.get("bullets", 0) for s in shapes)
    if total_bullets >= 5:
        return "content_slide"

    return "general_slide"
```

### 3. Infer Shape Purpose

#### Position-Based Inference

```python
def infer_shape_purpose_by_position(shape, slide_width, slide_height):
    """Infer purpose based on where shape appears on slide."""

    # Top 20% of slide
    if shape["position"]["top"] < slide_height * 0.2:
        if shape["width"] > slide_width * 0.6:
            return "slide_title"
        else:
            return "header_element"

    # Bottom 10% of slide
    elif shape["position"]["top"] > slide_height * 0.9:
        return "footer"

    # Left 30% of slide
    elif shape["position"]["left"] < slide_width * 0.3:
        return "sidebar_content"

    # Center mass
    elif (slide_width * 0.3 < shape["position"]["left"] < slide_width * 0.7):
        return "main_content"

    else:
        return "supplementary_content"
```

#### Content-Based Inference

```python
import re
from datetime import datetime

def infer_shape_purpose_by_content(text):
    """Infer purpose based on text patterns."""

    # Date patterns
    if re.search(r'\b(January|February|March|Q[1-4]|20\d{2})\b', text):
        return "date_field"

    # Metric patterns
    if re.search(r'^\d+%?$|^\$[\d,]+$', text.strip()):
        return "metric_value"

    # Header patterns (short, title case, < 60 chars)
    if len(text) < 60 and text.istitle() and '\n' not in text:
        return "header"

    # List content (bullets, multiple lines)
    if '\n' in text and len(text.split('\n')) >= 3:
        return "bullet_list"

    # Name/label patterns
    if text.endswith(':') or len(text.split()) <= 3:
        return "label"

    return "body_text"
```

### 4. Match Content to Shapes

Given new content (e.g., meeting transcript), map it to template shapes:

```python
def match_content_to_template(new_content, template_structure):
    """Map new content to appropriate template shapes."""

    updates = []

    # Extract structured data from new content
    extracted_data = extract_structured_data(new_content)

    for slide in template_structure["slides"]:
        for shape in slide["shapes"]:
            purpose = shape.get("inferred_purpose")

            # Match based on purpose
            if purpose == "date_field" and extracted_data.get("date"):
                updates.append({
                    "slide": slide["slide_number"],
                    "shape": shape["index"],
                    "text": extracted_data["date"]
                })

            elif purpose == "metric_value" and extracted_data.get("metrics"):
                # Match metrics by proximity to labels
                metric = find_closest_metric(shape, extracted_data["metrics"])
                if metric:
                    updates.append({
                        "slide": slide["slide_number"],
                        "shape": shape["index"],
                        "text": str(metric["value"])
                    })

            elif purpose == "bullet_list" and extracted_data.get("key_points"):
                # Format as bullets
                bullet_text = "\n".join(extracted_data["key_points"])
                updates.append({
                    "slide": slide["slide_number"],
                    "shape": shape["index"],
                    "text": bullet_text,
                    "preserve_bullets": True
                })

    return updates
```

## Semantic Patterns

### Pattern 1: Label-Value Pairs

Many templates have label-value pairs where labels are fixed and values change:

```
Revenue:  $2.5M
Users:    10,000
Growth:   25%
```

**Detection:**
```python
def find_label_value_pairs(shapes):
    """Find shapes that form label-value pairs."""
    pairs = []

    for i, shape in enumerate(shapes):
        if shape["text_content"].endswith(':'):
            # Look for nearby shapes (within 100 pixels to the right)
            for other in shapes:
                if (other["position"]["left"] > shape["position"]["left"] and
                    other["position"]["left"] - shape["position"]["left"] < 100 and
                    abs(other["position"]["top"] - shape["position"]["top"]) < 50):
                    pairs.append({
                        "label": shape,
                        "value": other
                    })

    return pairs
```

### Pattern 2: Hierarchical Bullets

Understand bullet hierarchy to preserve structure:

```
• Main Point 1
  - Sub point A
  - Sub point B
• Main Point 2
  - Sub point A
```

**Preservation:**
```python
def preserve_bullet_hierarchy(original_bullets, new_content):
    """Maintain bullet levels when updating."""

    # Analyze original structure
    original_levels = [p.level for p in original_bullets]
    has_hierarchy = len(set(original_levels)) > 1

    if has_hierarchy:
        # Try to maintain similar hierarchy in new content
        # E.g., first line = level 0, indented = level 1
        return format_with_hierarchy(new_content, original_levels)
    else:
        # Simple bullet list
        return new_content.split('\n')
```

### Pattern 3: Time-Based Content

Detect and update temporal references:

```
Q3 2024 Results
Week of Dec 16
January 2025 Goals
```

**Update Strategy:**
```python
def update_temporal_content(shape_text, new_date):
    """Replace old dates with new ones while preserving format."""

    # Detect format
    if re.match(r'Q[1-4] \d{4}', shape_text):
        quarter = (new_date.month - 1) // 3 + 1
        return f"Q{quarter} {new_date.year}"

    elif re.match(r'\w+ \d{4}', shape_text):
        return new_date.strftime('%B %Y')

    elif 'Week of' in shape_text:
        return f"Week of {new_date.strftime('%b %d')}"

    return new_date.strftime('%Y-%m-%d')
```

## Length Management

Critical for template updates: text must fit in shapes.

```python
def fit_text_to_shape(text, max_chars, preserve_meaning=True):
    """Truncate or summarize text to fit shape constraints."""

    if len(text) <= max_chars:
        return text

    if preserve_meaning:
        # Intelligent truncation
        # For bullets: remove less important points
        # For prose: summarize
        return summarize_to_length(text, max_chars)
    else:
        # Hard truncation with ellipsis
        return text[:max_chars-3] + "..."


def estimate_max_chars(shape_data):
    """Estimate safe character count based on shape size."""

    # Rough heuristic: area / character density
    area = shape_data["position"]["width"] * shape_data["position"]["height"]

    # Assume ~12pt font, ~10 chars per inch width
    # Adjust based on your templates
    estimated_chars = area // 10000  # EMUs to approximate chars

    # Add safety margin
    return int(estimated_chars * 0.8)
```

## Best Practices

1. **Always analyze before updating** - Run extraction first to understand structure
2. **Use multiple signals** - Combine position, content, and context for better accuracy
3. **Provide confidence scores** - Let user verify ambiguous matches
4. **Respect the original structure** - Don't add complexity the template doesn't have
5. **Test with real data** - Semantic inference improves with real-world examples
6. **Fall back gracefully** - When uncertain, ask the user which shape to update

## Example Workflow

```python
# 1. Extract template structure
structure = extract_template_structure("quarterly_report.pptx")

# 2. Analyze each shape's purpose
for slide in structure["slides"]:
    slide["type"] = infer_slide_type(slide)
    for shape in slide["shapes"]:
        shape["purpose"] = infer_shape_purpose(shape, slide)

# 3. Parse new content
new_data = parse_meeting_transcript("meeting_notes.txt")

# 4. Match content to shapes
updates = match_content_to_template(new_data, structure)

# 5. Validate lengths
for update in updates:
    shape = find_shape(structure, update["slide"], update["shape"])
    max_chars = estimate_max_chars(shape)
    if len(update["text"]) > max_chars:
        update["text"] = fit_text_to_shape(update["text"], max_chars)

# 6. Apply updates
apply_updates("quarterly_report.pptx", {"updates": updates}, "output.pptx")
```
