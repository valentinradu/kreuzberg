//! Unified rendering of `DocumentStructure` to output formats.
//!
//! Provides format-specific renderers that walk a document tree and produce
//! markdown, plain text, or HTML output. This replaces per-extractor ad-hoc
//! text generation with a single, consistent rendering pipeline.

mod markdown;
mod plain;

pub use markdown::render_to_markdown;
pub use plain::render_to_plain;
