//! Text formatting utilities for RTF content.

/// Normalize whitespace in a string.
///
/// - Collapses multiple consecutive spaces/tabs into a single space
/// - Preserves single newlines (paragraph breaks from \par)
/// - Collapses multiple consecutive newlines into a double newline
/// - Trims leading/trailing whitespace from each line
/// - Trims leading/trailing blank lines
pub fn normalize_whitespace(s: &str) -> String {
    // Split into lines, trim each, collapse blank runs
    let mut lines: Vec<&str> = Vec::new();
    let mut last_blank = false;

    for line in s.split('\n') {
        // Collapse internal whitespace on each line
        let trimmed = line.trim();
        if trimmed.is_empty() {
            if !last_blank && !lines.is_empty() {
                lines.push("");
                last_blank = true;
            }
        } else {
            last_blank = false;
            lines.push(trimmed);
        }
    }

    // Trim trailing blank lines
    while lines.last() == Some(&"") {
        lines.pop();
    }

    // Join and collapse internal multi-spaces within each line
    let joined = lines.join("\n");

    // Collapse runs of spaces within lines
    let mut result = String::with_capacity(joined.len());
    let mut last_was_space = false;
    for ch in joined.chars() {
        if ch == '\n' {
            result.push('\n');
            last_was_space = false;
        } else if ch == ' ' || ch == '\t' {
            if !last_was_space {
                result.push(' ');
                last_was_space = true;
            }
        } else {
            result.push(ch);
            last_was_space = false;
        }
    }

    // Remove spurious spaces before/after punctuation marks that result from RTF group boundaries
    let result = result.trim().to_string();
    let mut cleaned = String::with_capacity(result.len());
    let chars: Vec<char> = result.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == ' '
            && i + 1 < chars.len()
            && matches!(chars[i + 1], '.' | ',' | ';' | ':' | '!' | '?' | '|')
            && (i == 0 || chars[i - 1] != ' ')
        {
            // Skip the space before punctuation/pipe
            i += 1;
            continue;
        }
        if chars[i] == ' ' && i > 0 && chars[i - 1] == '|' {
            // Skip the space after pipe (table cell separator)
            i += 1;
            continue;
        }
        cleaned.push(chars[i]);
        i += 1;
    }
    cleaned
}
