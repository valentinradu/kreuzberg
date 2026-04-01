//! XML parsing and document structure traversal for JATS documents.

use crate::Result;
use crate::text::utf8_validation;
use quick_xml::Reader;
use quick_xml::events::Event;

/// Extract text content from a JATS element and its children.
pub(super) fn extract_text_content(reader: &mut Reader<&[u8]>) -> Result<String> {
    let mut text = String::new();
    let mut depth = 0;

    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                depth += 1;
            }
            Ok(Event::End(_)) => {
                if depth == 0 {
                    break;
                }
                depth -= 1;
                if !text.is_empty() && !text.ends_with('\n') {
                    text.push(' ');
                }
            }
            Ok(Event::Text(t)) => {
                let decoded = String::from_utf8_lossy(t.as_ref()).to_string();
                if !decoded.trim().is_empty() {
                    text.push_str(&decoded);
                    text.push(' ');
                }
            }
            Ok(Event::CData(t)) => {
                let decoded = utf8_validation::from_utf8(t.as_ref()).unwrap_or("").to_string();
                if !decoded.trim().is_empty() {
                    text.push_str(&decoded);
                    text.push('\n');
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(crate::error::KreuzbergError::parsing(format!(
                    "XML parsing error: {}",
                    e
                )));
            }
            _ => {}
        }
    }

    Ok(text.trim().to_string())
}
