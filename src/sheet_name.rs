use std::collections::HashSet;

/// Sanitizes a name for use as an Excel sheet name.
///
/// Excel sheet names have the following restrictions:
/// - Maximum 31 characters
/// - Cannot contain: \ / * ? : [ ]
/// - Cannot start or end with an apostrophe (')
/// - Cannot be named "History" (reserved word)
/// - Must be unique within the workbook
///
/// # Arguments
/// * `name` - The desired sheet name
/// * `existing` - A HashSet tracking already-used sheet names (will be updated)
///
/// # Returns
/// A sanitized sheet name that complies with Excel's naming rules
///
/// # Examples
/// ```
/// use std::collections::HashSet;
/// use sqlite_to_xlsx::sheet_name::sanitize_sheet_name;
///
/// let mut existing = HashSet::new();
/// let sheet_name = sanitize_sheet_name("My Sheet", &mut existing);
/// assert_eq!(sheet_name, "My Sheet");
/// assert!(existing.contains("My Sheet"));
/// ```
pub fn sanitize_sheet_name(name: &str, existing: &mut HashSet<String>) -> String {
    // Invalid characters for Excel sheet names
    const INVALID_CHARS: &[char] = &['\\', '/', '*', '?', ':', '[', ']'];
    const MAX_LENGTH: usize = 31;

    // Step 1: Replace invalid characters with underscores
    let sanitized: String = name
        .chars()
        .map(|c| if INVALID_CHARS.contains(&c) { '_' } else { c })
        .collect();

    // Step 2: Strip leading and trailing apostrophes
    let sanitized = sanitized.trim_matches('\'');

    // Step 3: Handle empty result
    let sanitized = if sanitized.is_empty() {
        "Sheet"
    } else {
        sanitized
    };

    // Step 4: Handle reserved name "History" (case-insensitive, preserve original case)
    let sanitized = if sanitized.eq_ignore_ascii_case("History") {
        format!("{}_", sanitized)
    } else {
        sanitized.to_string()
    };

    // Step 5: Truncate to 31 chars, leaving room for _N suffix if needed
    let base_name = if sanitized.chars().count() > MAX_LENGTH {
        // Truncate to 28 chars to leave room for _99 suffix (28 + 3 = 31)
        truncate_to_chars(&sanitized, 28)
    } else {
        sanitized
    };

    // Step 6: Ensure uniqueness by adding _1, _2, etc.
    let final_name = make_unique(&base_name, existing, MAX_LENGTH);

    // Step 7: Add to existing set
    existing.insert(final_name.clone());

    final_name
}

/// Makes a sheet name unique by appending a numeric suffix if necessary.
///
/// # Arguments
/// * `base_name` - The base name to make unique
/// * `existing` - HashSet of already-used names
///
/// # Returns
/// A unique name (either the base_name or base_name_N)
fn make_unique(base_name: &str, existing: &HashSet<String>, max_length: usize) -> String {
    // If the name is not taken, use it as-is
    if !existing.contains(base_name) {
        return base_name.to_string();
    }

    // Try _1, _2, _3, etc. until we find an unused name
    let mut counter = 1;
    loop {
        let suffix = format!("_{}", counter);
        let base_max = max_length.saturating_sub(suffix.chars().count());
        let trimmed_base = truncate_to_chars(base_name, base_max);
        let candidate = format!("{}{}", trimmed_base, suffix);
        if !existing.contains(&candidate) {
            return candidate;
        }
        counter += 1;

        // Safety check - prevent infinite loop in edge cases
        if counter > 10000 {
            // Generate a UUID-based name as fallback
            let suffix = format!("_{}", uuid_counter());
            let base_max = max_length.saturating_sub(suffix.chars().count());
            let trimmed_base = truncate_to_chars(base_name, base_max);
            return format!("{}{}", trimmed_base, suffix);
        }
    }
}

/// Generates a simple counter-based unique identifier as fallback.
fn uuid_counter() -> u32 {
    use std::sync::atomic::{AtomicU32, Ordering};
    static COUNTER: AtomicU32 = AtomicU32::new(1);
    COUNTER.fetch_add(1, Ordering::SeqCst)
}

fn truncate_to_chars(s: &str, max_chars: usize) -> String {
    if s.chars().count() <= max_chars {
        return s.to_string();
    }
    s.chars().take(max_chars).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_basic_valid_name() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet1", &mut existing);
        assert_eq!(result, "Sheet1");
        assert!(existing.contains("Sheet1"));
    }

    #[test]
    fn test_name_with_spaces() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("My Sheet", &mut existing);
        assert_eq!(result, "My Sheet");
        assert!(existing.contains("My Sheet"));
    }

    #[test]
    fn test_replace_backslash() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet\\Data", &mut existing);
        assert_eq!(result, "Sheet_Data");
        assert!(existing.contains("Sheet_Data"));
    }

    #[test]
    fn test_replace_forward_slash() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet/Data", &mut existing);
        assert_eq!(result, "Sheet_Data");
    }

    #[test]
    fn test_replace_asterisk() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet*Data", &mut existing);
        assert_eq!(result, "Sheet_Data");
    }

    #[test]
    fn test_replace_question_mark() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet?Data", &mut existing);
        assert_eq!(result, "Sheet_Data");
    }

    #[test]
    fn test_replace_colon() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet:Data", &mut existing);
        assert_eq!(result, "Sheet_Data");
    }

    #[test]
    fn test_replace_open_bracket() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet[Data", &mut existing);
        assert_eq!(result, "Sheet_Data");
    }

    #[test]
    fn test_replace_close_bracket() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet]Data", &mut existing);
        assert_eq!(result, "Sheet_Data");
    }

    #[test]
    fn test_multiple_invalid_chars() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet\\Data/Test:More[Info]", &mut existing);
        assert_eq!(result, "Sheet_Data_Test_More_Info_");
    }

    #[test]
    fn test_strip_leading_apostrophe() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("'Sheet1", &mut existing);
        assert_eq!(result, "Sheet1");
    }

    #[test]
    fn test_strip_trailing_apostrophe() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet1'", &mut existing);
        assert_eq!(result, "Sheet1");
    }

    #[test]
    fn test_strip_both_apostrophes() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("'Sheet1'", &mut existing);
        assert_eq!(result, "Sheet1");
    }

    #[test]
    fn test_only_apostrophes() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("'''", &mut existing);
        assert_eq!(result, "Sheet");
    }

    #[test]
    fn test_empty_string() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("", &mut existing);
        assert_eq!(result, "Sheet");
    }

    #[test]
    fn test_history_reserved_lowercase() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("History", &mut existing);
        assert_eq!(result, "History_");
    }

    #[test]
    fn test_history_reserved_uppercase() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("HISTORY", &mut existing);
        assert_eq!(result, "HISTORY_");
        assert!(existing.contains("HISTORY_"));
    }

    #[test]
    fn test_history_reserved_mixed_case() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("hIsToRy", &mut existing);
        assert_eq!(result, "hIsToRy_");
    }

    #[test]
    fn test_history_with_suffix_still_reserved() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("HistoryData", &mut existing);
        assert_eq!(result, "HistoryData");
    }

    #[test]
    fn test_long_name_truncation() {
        let mut existing = HashSet::new();
        let long_name = "ThisIsAVeryLongSheetNameThatExceedsTheLimit";
        let result = sanitize_sheet_name(long_name, &mut existing);
        // Should be truncated to 28 chars to leave room for _N suffix
        assert!(result.len() <= 31);
        assert_eq!(result, "ThisIsAVeryLongSheetNameThat");
        assert_eq!(result.len(), 28);
    }

    #[test]
    fn test_exactly_31_chars() {
        let mut existing = HashSet::new();
        let name = "1234567890123456789012345678901"; // 31 chars
        let result = sanitize_sheet_name(name, &mut existing);
        assert_eq!(result, name);
        assert_eq!(result.len(), 31);
    }

    #[test]
    fn test_one_over_limit() {
        let mut existing = HashSet::new();
        let name = "12345678901234567890123456789012"; // 32 chars
        let result = sanitize_sheet_name(name, &mut existing);
        assert_eq!(result.len(), 28);
    }

    #[test]
    fn test_long_name_with_invalid_chars() {
        let mut existing = HashSet::new();
        let long_name = "This/Is/A/Very/Long/Sheet/Name/With/Invalid/Chars/That/Exceeds/Limit";
        let result = sanitize_sheet_name(long_name, &mut existing);
        // After replacing / with _ and truncating
        assert!(result.len() <= 31);
    }

    #[test]
    fn test_single_duplicate() {
        let mut existing = HashSet::new();
        existing.insert("Sheet1".to_string());

        let result = sanitize_sheet_name("Sheet1", &mut existing);
        assert_eq!(result, "Sheet1_1");
        assert!(existing.contains("Sheet1_1"));
    }

    #[test]
    fn test_multiple_duplicates() {
        let mut existing = HashSet::new();
        existing.insert("Sheet1".to_string());
        existing.insert("Sheet1_1".to_string());
        existing.insert("Sheet1_2".to_string());

        let result = sanitize_sheet_name("Sheet1", &mut existing);
        assert_eq!(result, "Sheet1_3");
    }

    #[test]
    fn test_gaps_in_duplicate_numbers() {
        let mut existing = HashSet::new();
        existing.insert("Sheet1".to_string());
        existing.insert("Sheet1_1".to_string());
        existing.insert("Sheet1_5".to_string());

        let result = sanitize_sheet_name("Sheet1", &mut existing);
        // Should use the next available number (2, since we have 1 and 5)
        assert_eq!(result, "Sheet1_2");
    }

    #[test]
    fn test_first_name_no_duplicate() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("Sheet1", &mut existing);
        assert_eq!(result, "Sheet1");
        assert!(existing.contains("Sheet1"));
    }

    #[test]
    fn test_duplicate_then_another_duplicate() {
        let mut existing = HashSet::new();
        let result1 = sanitize_sheet_name("Table", &mut existing);
        assert_eq!(result1, "Table");

        let result2 = sanitize_sheet_name("Table", &mut existing);
        assert_eq!(result2, "Table_1");

        let result3 = sanitize_sheet_name("Table", &mut existing);
        assert_eq!(result3, "Table_2");
    }

    #[test]
    fn test_duplicate_with_sanitization() {
        let mut existing = HashSet::new();
        existing.insert("Sheet_Data".to_string());

        let result = sanitize_sheet_name("Sheet/Data", &mut existing);
        assert_eq!(result, "Sheet_Data_1");
    }

    #[test]
    fn test_history_duplicate() {
        let mut existing = HashSet::new();
        existing.insert("History_".to_string());

        let result = sanitize_sheet_name("History", &mut existing);
        assert_eq!(result, "History__1");
    }

    #[test]
    fn test_long_name_with_duplicate() {
        let mut existing = HashSet::new();
        let long_name = "ThisIsAVeryLongSheetNameThat";
        existing.insert(long_name.to_string());

        let result = sanitize_sheet_name(
            "ThisIsAVeryLongSheetNameThatExceedsTheLimit",
            &mut existing
        );
        assert!(result.len() <= 31);
        assert_eq!(result, "ThisIsAVeryLongSheetNameThat_1");
        assert_eq!(result.len(), 30); // 28 + _1 = 28 + 2 = 30
    }

    #[test]
    fn test_unicode_name() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("数据表", &mut existing);
        assert_eq!(result, "数据表");
    }

    #[test]
    fn test_unicode_with_invalid_chars() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("表/格", &mut existing);
        assert_eq!(result, "表_格");
    }

    #[test]
    fn test_only_invalid_chars() {
        let mut existing = HashSet::new();
        let result = sanitize_sheet_name("\\/*?:[]", &mut existing);
        // 7 invalid chars become 7 underscores, which is not empty, so it stays
        assert_eq!(result, "_______");
        assert_eq!(result.len(), 7);
    }

    #[test]
    fn test_name_stays_31_chars_with_duplicate() {
        let mut existing = HashSet::new();
        // Create a 28 char name (max base before suffix)
        let name_28 = "1234567890123456789012345678";
        existing.insert(name_28.to_string());

        let result = sanitize_sheet_name(name_28, &mut existing);
        // Result should be name_28_1 which is 30 chars total (28 + 2)
        assert_eq!(result, format!("{}_1", name_28));
        assert_eq!(result.len(), 30);
    }

    #[test]
    fn test_multiple_calls_consistent_tracking() {
        let mut existing = HashSet::new();

        let r1 = sanitize_sheet_name("Table", &mut existing);
        assert_eq!(r1, "Table");

        let r2 = sanitize_sheet_name("Table", &mut existing);
        assert_eq!(r2, "Table_1");

        let r3 = sanitize_sheet_name("Table", &mut existing);
        assert_eq!(r3, "Table_2");

        // Different base name should start fresh
        let r4 = sanitize_sheet_name("Other", &mut existing);
        assert_eq!(r4, "Other");

        let r5 = sanitize_sheet_name("Other", &mut existing);
        assert_eq!(r5, "Other_1");

        // Original sequence continues
        let r6 = sanitize_sheet_name("Table", &mut existing);
        assert_eq!(r6, "Table_3");
    }

    #[test]
    fn test_truncated_then_duplicated() {
        let mut existing = HashSet::new();
        let long = "A_Very_Long_Table_Name_Here_That_Exceeds";

        let r1 = sanitize_sheet_name(long, &mut existing);
        assert_eq!(r1, "A_Very_Long_Table_Name_Here_");
        assert_eq!(r1.len(), 28);

        let r2 = sanitize_sheet_name(long, &mut existing);
        // Should duplicate the truncated version
        assert_eq!(r2, "A_Very_Long_Table_Name_Here__1");
    }
}
