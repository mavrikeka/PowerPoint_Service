# Markdown Inline Formatting Implementation Plan

**Issue:** FUL-116
**Date:** 2026-02-19
**Author:** Vikram Ekambaram

## Executive Summary

Add support for markdown inline formatting (bold, italic, underline) to the PowerPoint generation service while maintaining 100% backward compatibility with existing plain-text usage.

## Problem Statement

The PPT generation service currently treats all text as plain strings. When AI agents provide markdown-formatted content (e.g., `**bold text**`, `*italic text*`), it renders as literal asterisks in PowerPoint instead of applying formatting.

### Current Behavior
```json
Input: "He is **immediately credible** on cost optimization"
Output in PPTX: "He is **immediately credible** on cost optimization" (literal asterisks)
```

### Expected Behavior
```json
Input: "He is **immediately credible** on cost optimization"
Output in PPTX: "He is immediately credible on cost optimization" (bold formatting applied)
```

## Solution Architecture

### Design Principles

1. **Zero Breaking Changes**: Plain text without markdown must work exactly as before
2. **Performance**: No overhead for plain text (fast-path detection)
3. **Simplicity**: Use regex-based parsing (no external dependencies)
4. **Robustness**: Handle edge cases gracefully (nested formatting, escaped characters, etc.)
5. **Consistency**: Apply template font properties to all formatted runs

### Technical Approach

#### Option Analysis

| Option | Pros | Cons | Decision |
|--------|------|------|----------|
| **Regex-based parser** | No dependencies, fast, simple | Manual parsing logic | ✅ **SELECTED** |
| `markdown` library | Full markdown support | Heavy, converts to HTML | ❌ Overkill |
| `mistune` library | Fast, lightweight | External dependency, full parser | ❌ Unnecessary complexity |
| `python-markdown` | Standard library | Too feature-rich | ❌ Not needed |

**Decision: Regex-based inline parser**
- Minimal code footprint
- No new dependencies
- Full control over behavior
- Optimal performance

## Implementation Specification

### 1. Supported Markdown Syntax (All 7 Combinations)

| Syntax | Meaning | Example | Output |
|--------|---------|---------|--------|
| `**text**` | Bold | `**strong**` | **strong** |
| `*text*` | Italic | `*emphasis*` | *emphasis* |
| `__text__` | Underline | `__important__` | <u>important</u> |
| `***text***` | Bold + Italic | `***both***` | ***both*** |
| `**__text__**` or `__**text**__` | Bold + Underline | `**__combined__**` | **<u>combined</u>** |
| `*__text__*` or `__*text*__` | Italic + Underline | `*__styled__*` | *<u>styled</u>* |
| `***__text__***` or `__***text***__` | Bold + Italic + Underline | `***__all__***` | ***<u>all</u>*** |

### 2. Edge Cases to Handle

#### A. Escaped Characters
```python
Input:  "Use \*asterisks\* for multiplication"
Output: "Use *asterisks* for multiplication" (literal asterisks, no formatting)
```

#### B. Unmatched Markers
```python
Input:  "This has **unmatched bold"
Output: "This has **unmatched bold" (literal asterisks, no formatting)
```

#### C. Whitespace Handling
```python
Input:  "** spaces around **"
Output: "** spaces around **" (no formatting - invalid markdown)

Input:  "**no spaces**"
Output: "no spaces" (bold applied)
```

#### D. Newlines in Formatted Text
```python
Input:  "**This spans\nmultiple lines**"
Output: Apply bold to entire run across newlines
```

#### E. Empty Formatted Sections
```python
Input:  "****"
Output: "" (empty bold section)
```

#### F. Multiple Formatting Types
```python
Input:  "**bold** and *italic* and __underline__"
Output: Three separate formatted runs with appropriate properties
```

### 3. Function Architecture

```
populate_text_placeholder(shape, text)
    ↓
    _has_markdown(text) → bool
    ↓
    If True: _parse_and_apply_markdown(text_frame, text, saved_font_props)
    If False: text_frame.text = text (existing fast path)
```

#### New Helper Functions

```python
def _has_markdown(text: str) -> bool:
    """
    Quick check if text contains markdown formatting.
    Returns True if text contains unescaped markdown markers.

    Pattern: **, *, __
    Skip escaped: \**, \*, \__
    """
    pass

def _parse_markdown_inline(text: str) -> List[Dict]:
    """
    Parse text into segments with formatting information.

    Returns:
        [
            {"text": "Plain text ", "bold": False, "italic": False, "underline": False},
            {"text": "bold text", "bold": True, "italic": False, "underline": False},
            {"text": " more plain", "bold": False, "italic": False, "underline": False}
        ]
    """
    pass

def _apply_markdown_to_paragraph(paragraph, segments, font_props):
    """
    Apply parsed markdown segments to a paragraph.
    Creates runs for each segment and applies formatting.

    Args:
        paragraph: pptx paragraph object
        segments: List of dicts from _parse_markdown_inline
        font_props: Font properties from template
    """
    pass

def _parse_and_apply_markdown(text_frame, text, saved_font_props):
    """
    Main markdown processing function.
    Handles multi-paragraph text (newlines) and applies formatting.
    """
    pass
```

### 4. Parsing Algorithm

#### Step 1: Pattern Matching (Regex)

```python
# Priority order (longest match first to handle nested formatting)
# Must match all 7 combinations in priority order
PATTERNS = [
    # All three (6 characters) - HIGHEST PRIORITY
    (r'(?<!\\)\*\*\*__(.+?)__\*\*\*', 'bold_italic_underline'),  # ***__text__***
    (r'(?<!\\)__\*\*\*(.+?)\*\*\*__', 'bold_italic_underline'),  # __***text***__

    # Two combinations (4-5 characters)
    (r'(?<!\\)\*\*__(.+?)__\*\*', 'bold_underline'),   # **__text__**
    (r'(?<!\\)__\*\*(.+?)\*\*__', 'bold_underline'),   # __**text**__
    (r'(?<!\\)\*__(.+?)__\*', 'italic_underline'),     # *__text__*
    (r'(?<!\\)__\*(.+?)\*__', 'italic_underline'),     # __*text*__
    (r'(?<!\\)\*\*\*(.+?)\*\*\*', 'bold_italic'),      # ***text***

    # Single formats (2-4 characters) - LOWEST PRIORITY
    (r'(?<!\\)\*\*(.+?)\*\*', 'bold'),                 # **text**
    (r'(?<!\\)__(.+?)__', 'underline'),                # __text__
    (r'(?<!\\)\*(.+?)\*', 'italic'),                   # *text*
]
```

#### Step 2: Segment Extraction

1. Find all markdown patterns in order of priority
2. Track start/end positions of each match
3. Build segments array with plain text + formatted text
4. Handle overlapping matches (first match wins)

#### Step 3: Run Creation

```python
for segment in segments:
    run = paragraph.add_run()
    run.text = segment['text']

    # Apply template font properties
    _apply_font_props(run.font, saved_font_props)

    # Apply markdown formatting
    if segment['bold']:
        run.font.bold = True
    if segment['italic']:
        run.font.italic = True
    if segment['underline']:
        run.font.underline = True
```

### 5. Backward Compatibility Strategy

#### Fast Path (No Markdown)
```python
if not _has_markdown(text):
    # Existing code path - zero overhead
    text_frame.text = text
    # Apply saved font props to all paragraphs
    for para in text_frame.paragraphs:
        if len(para.runs) > 0:
            _apply_font_props(para.runs[0].font, saved_font_props)
    return True
```

#### Markdown Path
```python
else:
    # New code path - only when markdown detected
    _parse_and_apply_markdown(text_frame, text, saved_font_props)
    return True
```

### 6. Testing Strategy

#### Unit Tests (Create: `test_markdown_formatting.py`)

```python
# Test cases:
1. Plain text (no markdown) - verify unchanged behavior
2. Simple bold: "**text**"
3. Simple italic: "*text*"
4. Simple underline: "__text__"
5. Bold + Italic: "***text***"
6. Bold + Underline: "**__text__**" and "__**text**__"
7. Italic + Underline: "*__text__*" and "__*text*__"
8. All three: "***__text__***" and "__***text***__"
9. Combined formatting: "**bold** and *italic* and __underline__"
10. Escaped characters: "\**not bold\**"
11. Unmatched markers: "**unclosed"
12. Empty formatted sections: "****"
13. Multiline text with markdown
14. Multiple paragraphs (newlines) with different formatting
15. Template font preservation across all 7 combinations
16. Unicode characters with markdown
17. Special characters: "Use * for multiplication"
18. Mixed combinations in one string: "**bold** ***bold italic*** *__italic underline__*"
```

#### Integration Tests

```python
# Test with actual PPTX templates:
1. Single-slide populate with markdown text
2. Multi-slide populate with mixed plain/markdown text
3. Table cells with markdown content
4. Font preservation across formatted runs
5. Performance test: 1000+ text fields (verify no slowdown)
```

### 7. Performance Considerations

#### Optimization 1: Early Exit
```python
def _has_markdown(text: str) -> bool:
    # Quick check - avoid regex if no markdown characters present
    if not any(c in text for c in ['*', '_']):
        return False
    # More thorough check only if needed
    return bool(re.search(r'(?<!\\)(\*\*|\*|__)', text))
```

#### Optimization 2: Compile Regex Patterns
```python
# Module-level constants - ALL 7 COMBINATIONS
MARKDOWN_PATTERNS = [
    # All three
    (re.compile(r'(?<!\\)\*\*\*__(.+?)__\*\*\*'), 'bold_italic_underline'),
    (re.compile(r'(?<!\\)__\*\*\*(.+?)\*\*\*__'), 'bold_italic_underline'),
    # Two combinations
    (re.compile(r'(?<!\\)\*\*__(.+?)__\*\*'), 'bold_underline'),
    (re.compile(r'(?<!\\)__\*\*(.+?)\*\*__'), 'bold_underline'),
    (re.compile(r'(?<!\\)\*__(.+?)__\*'), 'italic_underline'),
    (re.compile(r'(?<!\\)__\*(.+?)\*__'), 'italic_underline'),
    (re.compile(r'(?<!\\)\*\*\*(.+?)\*\*\*'), 'bold_italic'),
    # Single formats
    (re.compile(r'(?<!\\)\*\*(.+?)\*\*'), 'bold'),
    (re.compile(r'(?<!\\)__(.+?)__'), 'underline'),
    (re.compile(r'(?<!\\)\*(.+?)\*'), 'italic'),
]
```

#### Performance Targets
- Plain text: < 1ms overhead (fast path)
- Markdown text: < 10ms for typical paragraph (< 500 chars)
- No memory overhead for plain text

## Implementation Checklist

### Phase 1: Core Implementation
- [ ] Create `_has_markdown()` function
- [ ] Create `_parse_markdown_inline()` function
- [ ] Create `_apply_markdown_to_paragraph()` function
- [ ] Create `_parse_and_apply_markdown()` function
- [ ] Modify `populate_text_placeholder()` to use new functions
- [ ] Handle newlines (multiple paragraphs)

### Phase 2: Edge Case Handling
- [ ] Implement escape character handling (`\*`, `\_`)
- [ ] Handle unmatched markers gracefully
- [ ] Handle whitespace rules (no spaces around markers)
- [ ] Handle empty formatted sections
- [ ] Handle nested/combined formatting

### Phase 3: Testing
- [ ] Write unit tests for all markdown patterns
- [ ] Write unit tests for edge cases
- [ ] Write integration tests with real PPTX files
- [ ] Performance testing (plain text vs markdown)
- [ ] Backward compatibility validation

### Phase 4: Table Support
- [ ] Extend markdown support to `populate_table()` function
- [ ] Test table cells with markdown formatting
- [ ] Verify font preservation in table cells

### Phase 5: Documentation & Deployment
- [ ] Update API_DOCUMENTATION.md with markdown examples
- [ ] Update README.md with markdown formatting guide
- [ ] Add inline code comments
- [ ] Update CHANGELOG.md (if exists)
- [ ] Deploy to Railway (no new dependencies needed)

## Rollback Plan

If issues are discovered post-deployment:

1. **Quick Fix**: Add environment variable `DISABLE_MARKDOWN=true` to disable markdown parsing
2. **Code Rollback**: Revert to previous commit (markdown changes are isolated)
3. **Gradual Rollout**: Deploy to staging first, validate with real templates

## Success Criteria

1. ✅ All existing tests pass (zero regressions)
2. ✅ Plain text performance unchanged (< 1ms overhead)
3. ✅ Markdown formatting works for all supported patterns
4. ✅ Edge cases handled gracefully (no crashes)
5. ✅ Template font properties preserved
6. ✅ Works in both single-slide and multi-slide modes
7. ✅ Works in table cells

## Timeline Estimate

- Implementation: 2-3 hours
- Testing: 1-2 hours
- Documentation: 30 minutes
- **Total: 4-6 hours**

## Dependencies

- None (using only Python standard library + existing python-pptx)

## Risk Assessment

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| Breaking existing plain text | Low | High | Comprehensive testing, fast-path detection |
| Performance degradation | Low | Medium | Early exit optimization, benchmarking |
| Edge case bugs | Medium | Low | Extensive test coverage, graceful fallbacks |
| Regex complexity | Low | Low | Well-tested patterns, clear documentation |

## Future Enhancements (Out of Scope)

- Strikethrough: `~~text~~`
- Hyperlinks: `[text](url)`
- Code formatting: `` `code` ``
- Block-level markdown (headings, lists)
- Custom color/font via markdown
- Superscript/subscript formatting

---

## Approval

**Ready to Implement:** ✅
**Reviewed By:** Vikram Ekambaram
**Date:** 2026-02-19

This implementation plan covers all contingencies and provides a clear path forward with minimal risk.
