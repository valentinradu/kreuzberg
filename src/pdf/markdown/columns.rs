//! Column detection for multi-column PDF layouts.
//!
//! Detects column boundaries by analyzing the x-position distribution of
//! PDF page objects and splits them into separate column groups for
//! independent paragraph extraction.

use crate::pdf::hierarchy::SegmentData;
use pdfium_render::prelude::{PdfPageObject, PdfPageObjectCommon};

/// Minimum number of text objects per column to be considered valid.
const MIN_OBJECTS_PER_COLUMN: usize = 10;

/// Minimum gap between columns as fraction of page width.
const MIN_COLUMN_GAP_FRACTION: f32 = 0.04;

/// Minimum fraction of page height that both columns must span vertically.
const MIN_VERTICAL_SPAN_FRACTION: f32 = 0.3;

/// Minimum number of segments required to attempt a column split.
const MIN_SEGMENTS_FOR_SPLIT: usize = 10;

/// Minimum gap size between segments as fraction of the content span.
const SEGMENT_GAP_FRACTION: f32 = 0.05;

/// Maximum recursion depth for XY-Cut.
const MAX_XYCUT_DEPTH: usize = 4;

/// Split segments into column groups using recursive XY-Cut.
///
/// Returns a list of index-groups, each representing segments belonging to
/// the same column region, ordered left-to-right then top-to-bottom.
/// If no split is found, returns a single group with all indices.
pub(super) fn split_segments_into_columns(segments: &[SegmentData]) -> Vec<Vec<usize>> {
    let all_indices: Vec<usize> = (0..segments.len()).collect();
    xycut_recurse(segments, &all_indices, 0)
}

fn xycut_recurse(segments: &[SegmentData], indices: &[usize], depth: usize) -> Vec<Vec<usize>> {
    if indices.len() < MIN_SEGMENTS_FOR_SPLIT || depth >= MAX_XYCUT_DEPTH {
        return vec![indices.to_vec()];
    }

    // Compute bounding extent of these segments on x and y axes.
    let mut x_min = f32::MAX;
    let mut x_max = f32::MIN;
    let mut y_min = f32::MAX;
    let mut y_max = f32::MIN;
    for &i in indices {
        let s = &segments[i];
        let left = s.x;
        let right = s.x + s.width;
        let bottom = s.y;
        let top = s.y + s.height;
        x_min = x_min.min(left);
        x_max = x_max.max(right);
        y_min = y_min.min(bottom);
        y_max = y_max.max(top);
    }

    let x_span = x_max - x_min;
    let y_span = y_max - y_min;

    if x_span < 1.0 && y_span < 1.0 {
        return vec![indices.to_vec()];
    }

    // Try vertical cut (split left/right).
    let min_x_gap = x_span * SEGMENT_GAP_FRACTION;
    if let Some(split_x) = find_vertical_cut(segments, indices, min_x_gap, y_span) {
        let left: Vec<usize> = indices
            .iter()
            .copied()
            .filter(|&i| {
                let mid = segments[i].x + segments[i].width / 2.0;
                mid < split_x
            })
            .collect();
        let right: Vec<usize> = indices
            .iter()
            .copied()
            .filter(|&i| {
                let mid = segments[i].x + segments[i].width / 2.0;
                mid >= split_x
            })
            .collect();
        if !left.is_empty() && !right.is_empty() {
            let mut result = xycut_recurse(segments, &left, depth + 1);
            result.extend(xycut_recurse(segments, &right, depth + 1));
            return result;
        }
    }

    // Try horizontal cut (split top/bottom).
    let min_y_gap = y_span * SEGMENT_GAP_FRACTION;
    if let Some(split_y) = find_horizontal_cut(segments, indices, min_y_gap) {
        let top: Vec<usize> = indices
            .iter()
            .copied()
            .filter(|&i| {
                let mid = segments[i].y + segments[i].height / 2.0;
                mid >= split_y
            })
            .collect();
        let bottom: Vec<usize> = indices
            .iter()
            .copied()
            .filter(|&i| {
                let mid = segments[i].y + segments[i].height / 2.0;
                mid < split_y
            })
            .collect();
        if !top.is_empty() && !bottom.is_empty() {
            let mut result = xycut_recurse(segments, &top, depth + 1);
            result.extend(xycut_recurse(segments, &bottom, depth + 1));
            return result;
        }
    }

    vec![indices.to_vec()]
}

/// Find a vertical cut x-position by locating the largest horizontal gap.
///
/// Both sides of the cut must span at least `MIN_VERTICAL_SPAN_FRACTION` of `y_span`.
fn find_vertical_cut(segments: &[SegmentData], indices: &[usize], min_gap: f32, y_span: f32) -> Option<f32> {
    // Collect (left, right) edges sorted by left.
    let mut edges: Vec<(f32, f32)> = indices
        .iter()
        .map(|&i| (segments[i].x, segments[i].x + segments[i].width))
        .collect();
    edges.sort_by(|a, b| a.0.total_cmp(&b.0));

    let mut max_right = f32::MIN;
    let mut best_gap = 0.0_f32;
    let mut best_split: Option<f32> = None;

    for &(left, right) in &edges {
        if max_right > f32::MIN {
            let gap = left - max_right;
            if gap > min_gap && gap > best_gap {
                best_gap = gap;
                best_split = Some((max_right + left) / 2.0);
            }
        }
        max_right = max_right.max(right);
    }

    if let Some(split_x) = best_split
        && y_span >= 1.0
    {
        let left_y_span = vertical_span_of(segments, indices, |i| segments[i].x + segments[i].width / 2.0 < split_x);
        let right_y_span = vertical_span_of(segments, indices, |i| {
            segments[i].x + segments[i].width / 2.0 >= split_x
        });
        if left_y_span >= y_span * MIN_VERTICAL_SPAN_FRACTION && right_y_span >= y_span * MIN_VERTICAL_SPAN_FRACTION {
            return Some(split_x);
        }
    }

    None
}

/// Find a horizontal cut y-position by locating the largest vertical gap.
fn find_horizontal_cut(segments: &[SegmentData], indices: &[usize], min_gap: f32) -> Option<f32> {
    // Collect (bottom, top) edges sorted by bottom.
    let mut edges: Vec<(f32, f32)> = indices
        .iter()
        .map(|&i| (segments[i].y, segments[i].y + segments[i].height))
        .collect();
    edges.sort_by(|a, b| a.0.total_cmp(&b.0));

    let mut max_top = f32::MIN;
    let mut best_gap = 0.0_f32;
    let mut best_split: Option<f32> = None;

    for &(bottom, top) in &edges {
        if max_top > f32::MIN {
            let gap = bottom - max_top;
            if gap > min_gap && gap > best_gap {
                best_gap = gap;
                best_split = Some((max_top + bottom) / 2.0);
            }
        }
        max_top = max_top.max(top);
    }

    best_split
}

/// Compute vertical span (y extent) of a filtered subset of segments.
fn vertical_span_of<F>(segments: &[SegmentData], indices: &[usize], predicate: F) -> f32
where
    F: Fn(usize) -> bool,
{
    let mut y_min = f32::MAX;
    let mut y_max = f32::MIN;
    for &i in indices {
        if predicate(i) {
            let bottom = segments[i].y;
            let top = segments[i].y + segments[i].height;
            y_min = y_min.min(bottom);
            y_max = y_max.max(top);
        }
    }
    if y_max > y_min { y_max - y_min } else { 0.0 }
}

/// A bounding box extracted from a page object for column analysis.
struct ObjectBounds {
    left: f32,
    right: f32,
    top: f32,
    bottom: f32,
}

/// Detect column boundaries from page objects and return index groups.
///
/// Returns a list of index vectors, each representing objects belonging to
/// the same column, ordered left-to-right. If no columns are detected,
/// returns a single group containing all indices.
pub(super) fn split_objects_into_columns(objects: &[PdfPageObject]) -> Vec<Vec<usize>> {
    let bounds: Vec<ObjectBounds> = objects
        .iter()
        .filter_map(|obj| {
            // Only consider text objects for column detection
            obj.as_text_object()?;
            obj.bounds().ok().map(|b| ObjectBounds {
                left: b.left().value,
                right: b.right().value,
                top: b.top().value,
                bottom: b.bottom().value,
            })
        })
        .collect();

    if bounds.len() < MIN_OBJECTS_PER_COLUMN * 2 {
        return vec![(0..objects.len()).collect()];
    }

    let (page_width, page_y_min, page_y_max) = estimate_page_bounds(&bounds);
    if page_width < 1.0 {
        return vec![(0..objects.len()).collect()];
    }

    let min_gap = page_width * MIN_COLUMN_GAP_FRACTION;

    if let Some(split_x) = find_column_split(&bounds, min_gap, page_y_min, page_y_max) {
        let mut left_indices: Vec<usize> = Vec::new();
        let mut right_indices: Vec<usize> = Vec::new();

        // Partition ALL objects (not just text) by midpoint relative to split
        for (i, obj) in objects.iter().enumerate() {
            let mid_x = obj
                .bounds()
                .ok()
                .map(|b| (b.left().value + b.right().value) / 2.0)
                .unwrap_or(0.0);

            if mid_x < split_x {
                left_indices.push(i);
            } else {
                right_indices.push(i);
            }
        }

        // Validate column sizes (text objects only)
        let left_text_count = left_indices
            .iter()
            .filter(|&&i| objects[i].as_text_object().is_some())
            .count();
        let right_text_count = right_indices
            .iter()
            .filter(|&&i| objects[i].as_text_object().is_some())
            .count();

        if left_text_count < MIN_OBJECTS_PER_COLUMN || right_text_count < MIN_OBJECTS_PER_COLUMN {
            return vec![(0..objects.len()).collect()];
        }

        vec![left_indices, right_indices]
    } else {
        vec![(0..objects.len()).collect()]
    }
}

/// Estimate page bounds from object bounding boxes.
fn estimate_page_bounds(bounds: &[ObjectBounds]) -> (f32, f32, f32) {
    let mut x_min = f32::MAX;
    let mut x_max = f32::MIN;
    let mut y_min = f32::MAX;
    let mut y_max = f32::MIN;

    for b in bounds {
        x_min = x_min.min(b.left);
        x_max = x_max.max(b.right);
        y_min = y_min.min(b.bottom);
        y_max = y_max.max(b.top);
    }

    (x_max - x_min, y_min, y_max)
}

/// Find the best x-position to split columns using gap analysis.
///
/// Sorts object left edges, finds the widest gap exceeding `min_gap`,
/// and validates that objects on both sides span enough of the page height.
fn find_column_split(bounds: &[ObjectBounds], min_gap: f32, page_y_min: f32, page_y_max: f32) -> Option<f32> {
    let page_y_range = page_y_max - page_y_min;
    if page_y_range < 1.0 {
        return None;
    }

    // Collect (left, right) edges sorted by left edge
    let mut edges: Vec<(f32, f32)> = bounds.iter().map(|b| (b.left, b.right)).collect();
    edges.sort_by(|a, b| a.0.total_cmp(&b.0));

    // Track the running maximum right edge to find true gaps
    let mut max_right = f32::MIN;
    let mut best_gap = 0.0_f32;
    let mut best_split = None;

    for &(left, right) in &edges {
        if max_right > f32::MIN {
            let gap = left - max_right;
            if gap > min_gap && gap > best_gap {
                best_gap = gap;
                best_split = Some((max_right + left) / 2.0);
            }
        }
        max_right = max_right.max(right);
    }

    // Validate: both sides must span a significant portion of page height
    if let Some(split_x) = best_split {
        let left_y_range = vertical_span(bounds.iter().filter(|b| b.left < split_x));
        let right_y_range = vertical_span(bounds.iter().filter(|b| b.left >= split_x));

        if left_y_range > page_y_range * MIN_VERTICAL_SPAN_FRACTION
            && right_y_range > page_y_range * MIN_VERTICAL_SPAN_FRACTION
        {
            return Some(split_x);
        }
    }

    None
}

/// Compute the vertical span (top - bottom) of an iterator of bounds.
fn vertical_span<'a>(bounds: impl Iterator<Item = &'a ObjectBounds>) -> f32 {
    let mut y_min = f32::MAX;
    let mut y_max = f32::MIN;

    for b in bounds {
        y_min = y_min.min(b.bottom);
        y_max = y_max.max(b.top);
    }

    if y_max > y_min { y_max - y_min } else { 0.0 }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_empty_returns_single_group() {
        let objects: Vec<PdfPageObject> = vec![];
        let groups = split_objects_into_columns(&objects);
        assert_eq!(groups.len(), 1);
        assert!(groups[0].is_empty());
    }

    fn make_segment(x: f32, y: f32, w: f32, h: f32) -> SegmentData {
        SegmentData {
            text: "word".to_string(),
            x,
            y,
            width: w,
            height: h,
            font_size: 12.0,
            is_bold: false,
            is_italic: false,
            is_monospace: false,
            baseline_y: y,
        }
    }

    fn make_column_segments(x_offset: f32, count: usize) -> Vec<SegmentData> {
        (0..count)
            .map(|i| make_segment(x_offset, i as f32 * 20.0, 80.0, 12.0))
            .collect()
    }

    #[test]
    fn test_split_segments_too_few_returns_single_group() {
        let segments: Vec<SegmentData> = (0..5).map(|i| make_segment(i as f32 * 10.0, 0.0, 8.0, 12.0)).collect();
        let groups = split_segments_into_columns(&segments);
        assert_eq!(groups.len(), 1);
        assert_eq!(groups[0].len(), 5);
    }

    #[test]
    fn test_split_segments_empty_returns_single_group() {
        let segments: Vec<SegmentData> = vec![];
        let groups = split_segments_into_columns(&segments);
        assert_eq!(groups.len(), 1);
        assert!(groups[0].is_empty());
    }

    #[test]
    fn test_split_segments_two_columns_detected() {
        // Left column: x=0..80, right column: x=300..380, large gap in between.
        let mut segments = make_column_segments(0.0, 15);
        segments.extend(make_column_segments(300.0, 15));
        let groups = split_segments_into_columns(&segments);
        assert_eq!(groups.len(), 2, "expected 2 column groups, got {:?}", groups.len());
        // Each group should have 15 segments.
        assert_eq!(groups[0].len(), 15);
        assert_eq!(groups[1].len(), 15);
    }

    #[test]
    fn test_split_segments_single_column_no_false_split() {
        // All segments in a tight horizontal band — no real gap.
        let segments: Vec<SegmentData> = (0..20).map(|i| make_segment(i as f32 * 10.0, 0.0, 8.0, 12.0)).collect();
        let groups = split_segments_into_columns(&segments);
        assert_eq!(groups.len(), 1);
        assert_eq!(groups[0].len(), 20);
    }

    #[test]
    fn test_split_segments_indices_cover_all() {
        let mut segments = make_column_segments(0.0, 12);
        segments.extend(make_column_segments(300.0, 12));
        let groups = split_segments_into_columns(&segments);
        let total: usize = groups.iter().map(|g| g.len()).sum();
        assert_eq!(total, segments.len(), "all segment indices must be accounted for");
    }

    #[test]
    fn test_split_segments_depth_limit_prevents_over_segmentation() {
        // Many tiny columns that would over-segment without depth limit.
        let mut segments = Vec::new();
        for col in 0..10usize {
            for row in 0..5usize {
                segments.push(make_segment(col as f32 * 50.0, row as f32 * 20.0, 10.0, 12.0));
            }
        }
        let groups = split_segments_into_columns(&segments);
        // Depth limit of 4 means at most 2^4=16 groups, but content doesn't have enough
        // segments per group at deep levels, so it should be reasonable.
        assert!(groups.len() <= 16, "too many groups: {}", groups.len());
        let total: usize = groups.iter().map(|g| g.len()).sum();
        assert_eq!(total, segments.len());
    }
}
