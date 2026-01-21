# Python-PPTX Quick Reference

Key patterns for reading and writing PowerPoint files using python-pptx.

## Installation

```bash
pip install python-pptx
```

## Opening and Saving Presentations

```python
from pptx import Presentation

# Open existing presentation
prs = Presentation('template.pptx')

# Create new presentation
prs = Presentation()

# Save presentation
prs.save('output.pptx')
```

## Iterating Through Slides and Shapes

```python
# Iterate through slides
for slide_idx, slide in enumerate(prs.slides):
    print(f"Slide {slide_idx + 1}")

    # Iterate through shapes on slide
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            print(f"  Shape: {shape.name} - {shape.text}")
```

## Working with Text

### Reading Text from Shapes

```python
# Check if shape has text
if hasattr(shape, "text_frame") and shape.has_text_frame:
    text = shape.text_frame.text
    print(text)
```

### Updating Text (Simple)

```python
# Simple text replacement (loses formatting)
shape.text_frame.text = "New text"
```

### Updating Text (Preserving Structure)

```python
# Preserve paragraph structure
text_frame = shape.text_frame

# Clear existing text while keeping paragraphs
for paragraph in text_frame.paragraphs:
    paragraph.clear()

# Update first paragraph
text_frame.paragraphs[0].text = "First line"

# Add more paragraphs
p = text_frame.add_paragraph()
p.text = "Second line"
p.level = 1  # Indent level (0 = no indent, 1 = first level, etc.)
```

## Working with Paragraphs and Runs

```python
# Access paragraphs in a text frame
for paragraph in shape.text_frame.paragraphs:
    print(f"Level: {paragraph.level}")
    print(f"Text: {paragraph.text}")

    # Access runs (formatted text segments) in paragraph
    for run in paragraph.runs:
        print(f"  Run text: {run.text}")
        print(f"  Bold: {run.font.bold}")
        print(f"  Size: {run.font.size}")
```

## Preserving Formatting When Updating

### Method 1: Update run-by-run

```python
# Preserve formatting by updating each run
for run in paragraph.runs:
    # Store formatting
    font = run.font

    # Update text
    run.text = "New text"

    # Formatting is automatically preserved
```

### Method 2: Clear and recreate with same format

```python
# Store the first run's format
if paragraph.runs:
    original_font = paragraph.runs[0].font
    original_bold = original_font.bold
    original_size = original_font.size

# Clear paragraph
paragraph.clear()

# Add new text with preserved formatting
run = paragraph.add_run()
run.text = "New text"
run.font.bold = original_bold
run.font.size = original_size
```

## Checking Text Length Constraints

```python
# Get shape dimensions
width = shape.width
height = shape.height

# Get current text length
current_text = shape.text_frame.text
current_length = len(current_text)

# Warn if new text is significantly longer
new_length = len(new_text)
if new_length > current_length * 1.5:
    print(f"Warning: New text ({new_length} chars) is 50% longer than original ({current_length} chars)")
```

## Working with Placeholders

```python
# Check if shape is a placeholder
if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
    placeholder_type = shape.placeholder_format.type
    print(f"Placeholder type: {placeholder_type}")
```

## Shape Types

```python
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Check shape type
if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
    print("This is a text box")
elif shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
    print("This is a placeholder")
elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
    print("This is a table")
elif shape.shape_type == MSO_SHAPE_TYPE.CHART:
    print("This is a chart")
```

## Shape Position and Size

```python
from pptx.util import Inches, Pt

# Get position (in EMUs - English Metric Units)
left = shape.left
top = shape.top
width = shape.width
height = shape.height

# Convert to inches
left_inches = Inches(left)

# Convert from inches to EMUs
shape.left = Inches(1.5)
shape.width = Inches(3.0)
```

## Common Pitfalls

### 1. Text Frame vs Text

```python
# ❌ Wrong: Not all shapes have .text
text = shape.text  # May cause AttributeError

# ✅ Correct: Check for text_frame first
if hasattr(shape, "text_frame") and shape.has_text_frame:
    text = shape.text_frame.text
```

### 2. Losing Formatting

```python
# ❌ Wrong: Loses all formatting
shape.text_frame.text = "New text"

# ✅ Correct: Update paragraphs or runs to preserve structure
for paragraph in shape.text_frame.paragraphs:
    paragraph.text = "New text"
```

### 3. Not Handling Bullet Points

```python
# ❌ Wrong: Loses bullet structure
shape.text = "Line 1\nLine 2\nLine 3"

# ✅ Correct: Create separate paragraphs
text_frame = shape.text_frame
text_frame.clear()

for line in ["Line 1", "Line 2", "Line 3"]:
    p = text_frame.add_paragraph()
    p.text = line
    p.level = 1  # Maintain bullet level
```

## Best Practices for Template Updates

1. **Always preserve structure**: Update paragraphs/runs individually rather than replacing entire text frames
2. **Check length**: Warn if new text significantly exceeds original length
3. **Maintain bullet levels**: Copy `paragraph.level` when creating new paragraphs
4. **Test with actual templates**: Different templates may have different formatting quirks
5. **Handle errors gracefully**: Some shapes may not have text frames or may be grouped
