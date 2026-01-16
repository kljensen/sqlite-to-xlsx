use anyhow::Result;
use rust_xlsxwriter::Worksheet;
use rusqlite::types::Value;
use std::fmt::Write;

/// Configuration for how BLOB values should be written to Excel cells
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BlobHandling {
    /// Write a placeholder string like "[BLOB: 123 bytes]"
    Placeholder,
    /// Write hexadecimal representation like "0x48656c6c6f"
    Hex,
    /// Write base64 encoded string
    Base64,
    /// Write an empty cell (skip the BLOB value)
    Skip,
}

impl Default for BlobHandling {
    fn default() -> Self {
        BlobHandling::Placeholder
    }
}

/// Maximum safe integer for f64 (2^53)
/// Numbers larger than this cannot be precisely represented in f64
const MAX_SAFE_INTEGER: i64 = 9_007_199_254_740_992; // 2^53

/// Maximum string length before truncation (Excel has a 32,767 char limit)
const MAX_STRING_LENGTH: usize = 32_767;

/// Truncation suffix for overly long strings
const TRUNCATION_SUFFIX: &str = "... [truncated]";

/// Converts a SQLite value and writes it to an Excel worksheet cell
///
/// # Arguments
/// * `worksheet` - The Excel worksheet to write to
/// * `row` - The row index (0-based)
/// * `col` - The column index (0-based)
/// * `value` - The SQLite value to convert
/// * `blob_handling` - How to handle BLOB values
///
/// # Type Mapping
/// - INTEGER (i64) -> Number via write_number()
/// - REAL (f64) -> Number via write_number()
/// - TEXT (String) -> String via write_string()
/// - BLOB (Vec<u8>) -> Based on BlobHandling configuration
/// - NULL -> Blank via write_blank()
///
/// # Edge Cases
/// - NaN/Infinity floats are written as strings
/// - Large integers (>2^53) trigger a warning and are written as strings
/// - Very long strings (>32K chars) are truncated with a warning
pub fn write_cell(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &Value,
    blob_handling: BlobHandling,
) -> Result<()> {
    match value {
        Value::Null => {
            worksheet.write(row, col, None::<&str>).map_err(|e| anyhow::anyhow!(e))?;
        }
        Value::Integer(int_val) => {
            write_integer(worksheet, row, col, *int_val)?;
        }
        Value::Real(real_val) => {
            write_real(worksheet, row, col, *real_val)?;
        }
        Value::Text(text_val) => {
            write_string(worksheet, row, col, text_val)?;
        }
        Value::Blob(blob_val) => {
            write_blob(worksheet, row, col, blob_val, blob_handling)?;
        }
    }
    Ok(())
}

/// Writes an integer value to a worksheet cell
///
/// For large integers that exceed f64 precision (2^53), writes as string
/// to preserve precision and logs a warning.
fn write_integer(worksheet: &mut Worksheet, row: u32, col: u16, value: i64) -> Result<()> {
    if value.abs() > MAX_SAFE_INTEGER {
        eprintln!(
            "Warning: Large integer {} at row {}, col {} exceeds f64 precision, writing as string",
            value, row, col
        );
        worksheet
            .write_string(row, col, &value.to_string())
            .map_err(|e| anyhow::anyhow!(e))?;
    } else {
        worksheet
            .write_number(row, col, value as f64)
            .map_err(|e| anyhow::anyhow!(e))?;
    }
    Ok(())
}

/// Writes a real (float) value to a worksheet cell
///
/// Special float values (NaN, Infinity, -Infinity) are written as strings
/// since Excel cannot represent them as numbers.
fn write_real(worksheet: &mut Worksheet, row: u32, col: u16, value: f64) -> Result<()> {
    if value.is_nan() {
        worksheet
            .write_string(row, col, "NaN")
            .map_err(|e| anyhow::anyhow!(e))?;
    } else if value.is_infinite() {
        if value.is_sign_positive() {
            worksheet
                .write_string(row, col, "Infinity")
                .map_err(|e| anyhow::anyhow!(e))?;
        } else {
            worksheet
                .write_string(row, col, "-Infinity")
                .map_err(|e| anyhow::anyhow!(e))?;
        }
    } else {
        worksheet
            .write_number(row, col, value)
            .map_err(|e| anyhow::anyhow!(e))?;
    }
    Ok(())
}

/// Writes a string value to a worksheet cell
///
/// Truncates strings that exceed Excel's maximum length (32,767 characters)
/// and logs a warning.
fn write_string(worksheet: &mut Worksheet, row: u32, col: u16, value: &str) -> Result<()> {
    if value.len() > MAX_STRING_LENGTH {
        eprintln!(
            "Warning: String at row {}, col {} exceeds Excel max length ({} chars), truncating",
            row, col, value.len()
        );
        let truncated = truncate_string(value, MAX_STRING_LENGTH);
        worksheet
            .write_string(row, col, &truncated)
            .map_err(|e| anyhow::anyhow!(e))?;
    } else {
        worksheet
            .write_string(row, col, value)
            .map_err(|e| anyhow::anyhow!(e))?;
    }
    Ok(())
}

/// Writes a BLOB value to a worksheet cell based on the specified handling strategy
fn write_blob(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &[u8],
    blob_handling: BlobHandling,
) -> Result<()> {
    match blob_handling {
        BlobHandling::Placeholder => {
            let placeholder = format!("[BLOB: {} bytes]", value.len());
            worksheet
                .write_string(row, col, &placeholder)
                .map_err(|e| anyhow::anyhow!(e))?;
        }
        BlobHandling::Hex => {
            let hex_string = bytes_to_hex(value);
            write_string(worksheet, row, col, &hex_string)?;
        }
        BlobHandling::Base64 => {
            let base64_string = base64_encode(value);
            write_string(worksheet, row, col, &base64_string)?;
        }
        BlobHandling::Skip => {
            worksheet.write(row, col, None::<&str>).map_err(|e| anyhow::anyhow!(e))?;
        }
    }
    Ok(())
}

/// Truncates a string to fit within the maximum length
///
/// Preserves room for the truncation suffix indicator.
fn truncate_string(s: &str, max_len: usize) -> String {
    // If string fits within max_len, return it as-is
    if s.len() <= max_len {
        return s.to_string();
    }
    // If max_len is too small for the suffix, return just the suffix
    if max_len <= TRUNCATION_SUFFIX.len() {
        return TRUNCATION_SUFFIX.to_string();
    }
    // Truncate the string and add suffix
    let available = max_len - TRUNCATION_SUFFIX.len();
    let mut result = String::with_capacity(max_len);
    for c in s.chars().take(available) {
        result.push(c);
    }
    result.push_str(TRUNCATION_SUFFIX);
    result
}

/// Converts bytes to hexadecimal string representation
fn bytes_to_hex(bytes: &[u8]) -> String {
    let mut hex_string = String::with_capacity(bytes.len() * 2 + 2);
    hex_string.push_str("0x");
    for byte in bytes {
        write!(hex_string, "{:02x}", byte).expect("writing to string should not fail");
    }
    hex_string
}

/// Base64 encodes a byte slice
fn base64_encode(bytes: &[u8]) -> String {
    use base64::prelude::*;
    BASE64_STANDARD.encode(bytes)
}

/// Converts SQLite database to XLSX format
pub fn convert() -> Result<()> {
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use rust_xlsxwriter::Workbook;

    #[test]
    fn test_blob_handling_default() {
        let default = BlobHandling::default();
        assert_eq!(default, BlobHandling::Placeholder);
    }

    #[test]
    fn test_write_null() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let result = write_cell(worksheet, 0, 0, &Value::Null, BlobHandling::default());
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_integer() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let result = write_cell(worksheet, 0, 0, &Value::Integer(42), BlobHandling::default());
        assert!(result.is_ok());

        let result = write_cell(worksheet, 0, 1, &Value::Integer(-100), BlobHandling::default());
        assert!(result.is_ok());

        let result = write_cell(worksheet, 0, 2, &Value::Integer(0), BlobHandling::default());
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_large_integer() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let large_int = 10_000_000_000_000_000_i64; // Larger than 2^53
        let result = write_cell(worksheet, 0, 0, &Value::Integer(large_int), BlobHandling::default());
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_real() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let result = write_cell(worksheet, 0, 0, &Value::Real(3.14159), BlobHandling::default());
        assert!(result.is_ok());

        let result = write_cell(worksheet, 0, 1, &Value::Real(-2.5), BlobHandling::default());
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_nan() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let result = write_cell(worksheet, 0, 0, &Value::Real(f64::NAN), BlobHandling::default());
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_infinity() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let result = write_cell(worksheet, 0, 0, &Value::Real(f64::INFINITY), BlobHandling::default());
        assert!(result.is_ok());

        let result = write_cell(worksheet, 0, 1, &Value::Real(f64::NEG_INFINITY), BlobHandling::default());
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_text() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let result = write_cell(
            worksheet,
            0,
            0,
            &Value::Text("Hello, World!".to_string()),
            BlobHandling::default(),
        );
        assert!(result.is_ok());

        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        let result = write_cell(worksheet, 0, 0, &Value::Text("".to_string()), BlobHandling::default());
        assert!(result.is_ok());

        // Test Unicode
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        let result = write_cell(
            worksheet,
            0,
            0,
            &Value::Text("你好世界 🌍".to_string()),
            BlobHandling::default(),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_blob_placeholder() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let blob = vec![0x48, 0x65, 0x6c, 0x6c, 0x6f]; // "Hello"
        let result = write_cell(worksheet, 0, 0, &Value::Blob(blob), BlobHandling::Placeholder);
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_blob_hex() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let blob = vec![0xDE, 0xAD, 0xBE, 0xEF];
        let result = write_cell(worksheet, 0, 0, &Value::Blob(blob), BlobHandling::Hex);
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_blob_base64() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let blob = vec![0x48, 0x65, 0x6c, 0x6c, 0x6f]; // "Hello"
        let result = write_cell(worksheet, 0, 0, &Value::Blob(blob), BlobHandling::Base64);
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_blob_skip() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let blob = vec![0x01, 0x02, 0x03, 0x04];
        let result = write_cell(worksheet, 0, 0, &Value::Blob(blob), BlobHandling::Skip);
        assert!(result.is_ok());
    }

    #[test]
    fn test_write_empty_blob() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let blob = vec![];
        let result = write_cell(worksheet, 0, 0, &Value::Blob(blob), BlobHandling::Placeholder);
        assert!(result.is_ok());
    }

    #[test]
    fn test_bytes_to_hex() {
        assert_eq!(bytes_to_hex(&[]), "0x");
        assert_eq!(bytes_to_hex(&[0x00]), "0x00");
        assert_eq!(bytes_to_hex(&[0xFF]), "0xff");
        assert_eq!(bytes_to_hex(&[0xDE, 0xAD, 0xBE, 0xEF]), "0xdeadbeef");
        assert_eq!(bytes_to_hex(&[0x01, 0x02, 0x03]), "0x010203");
    }

    #[test]
    fn test_base64_encode() {
        assert_eq!(base64_encode(&[]), "");
        assert_eq!(base64_encode(&[0x48]), "SA==");
        assert_eq!(base64_encode(&[0x48, 0x65]), "SGU=");
        assert_eq!(base64_encode(&[0x48, 0x65, 0x6c, 0x6c, 0x6f]), "SGVsbG8=");
        assert_eq!(base64_encode(b"Hello, World!"), "SGVsbG8sIFdvcmxkIQ==");
    }

    #[test]
    fn test_truncate_string() {
        // Short strings should not be truncated
        let short = "Hello";
        assert_eq!(truncate_string(short, 100), "Hello");

        // Exactly at limit should not be truncated
        let exact = "a".repeat(50);
        assert_eq!(truncate_string(&exact, 50), exact);

        // Over limit should be truncated
        let long = "a".repeat(100);
        let truncated = truncate_string(&long, 50);
        assert_eq!(truncated.len(), 50);
        assert!(truncated.ends_with(TRUNCATION_SUFFIX));

        // Very short limit
        let very_short = truncate_string("Hello, World!", 5);
        assert_eq!(very_short, TRUNCATION_SUFFIX);
    }

    #[test]
    fn test_blob_handling_equality() {
        assert_eq!(BlobHandling::Placeholder, BlobHandling::Placeholder);
        assert_eq!(BlobHandling::Hex, BlobHandling::Hex);
        assert_eq!(BlobHandling::Base64, BlobHandling::Base64);
        assert_eq!(BlobHandling::Skip, BlobHandling::Skip);

        assert_ne!(BlobHandling::Placeholder, BlobHandling::Hex);
        assert_ne!(BlobHandling::Hex, BlobHandling::Base64);
        assert_ne!(BlobHandling::Base64, BlobHandling::Skip);
    }

    #[test]
    fn test_multiple_cells_same_row() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let values = vec![
            Value::Null,
            Value::Integer(42),
            Value::Real(3.14),
            Value::Text("test".to_string()),
            Value::Blob(vec![1, 2, 3]),
        ];

        for (col, value) in values.iter().enumerate() {
            let result = write_cell(worksheet, 0, col as u16, value, BlobHandling::default());
            assert!(result.is_ok());
        }
    }

    #[test]
    fn test_multiple_cells_same_column() {
        let mut workbook = Workbook::new();
        let mut worksheet = workbook.add_worksheet();

        let values = vec![
            Value::Integer(1),
            Value::Integer(2),
            Value::Integer(3),
        ];

        for (row, value) in values.iter().enumerate() {
            let result = write_cell(&mut worksheet, row as u32, 0, value, BlobHandling::default());
            assert!(result.is_ok());
        }
    }

    #[test]
    fn test_all_blob_handling_strategies() {
        let blob = vec![0x00, 0x01, 0x02, 0x03];

        for strategy in &[BlobHandling::Placeholder, BlobHandling::Hex, BlobHandling::Base64, BlobHandling::Skip] {
            let mut workbook = Workbook::new();
            let worksheet = workbook.add_worksheet();

            let result = write_cell(worksheet, 0, 0, &Value::Blob(blob.clone()), *strategy);
            assert!(result.is_ok(), "Failed for strategy: {:?}", strategy);
        }
    }

    #[test]
    fn test_special_float_values() {
        let mut workbook = Workbook::new();
        let mut worksheet = workbook.add_worksheet();

        // Test positive and negative zero (should both work)
        let result = write_cell(&mut worksheet, 0, 0, &Value::Real(0.0), BlobHandling::default());
        assert!(result.is_ok());

        let result = write_cell(&mut worksheet, 0, 1, &Value::Real(-0.0), BlobHandling::default());
        assert!(result.is_ok());

        // Test subnormal numbers
        let result = write_cell(&mut worksheet, 0, 2, &Value::Real(1e-320), BlobHandling::default());
        assert!(result.is_ok());
    }

    #[test]
    fn test_empty_string() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let result = write_cell(worksheet, 0, 0, &Value::Text(String::new()), BlobHandling::default());
        assert!(result.is_ok());
    }

    #[test]
    fn test_string_with_newlines() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let text = "Line 1\nLine 2\nLine 3";
        let result = write_cell(worksheet, 0, 0, &Value::Text(text.to_string()), BlobHandling::default());
        assert!(result.is_ok());
    }

    #[test]
    fn test_string_with_quotes() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let text = "He said \"Hello\"";
        let result = write_cell(worksheet, 0, 0, &Value::Text(text.to_string()), BlobHandling::default());
        assert!(result.is_ok());
    }

    #[test]
    fn test_large_blob_sizes() {
        let small_blob = vec![0u8; 10];
        let medium_blob = vec![0u8; 1_000];
        let large_blob = vec![0u8; 100_000];

        for blob in [small_blob, medium_blob, large_blob].iter() {
            let mut workbook = Workbook::new();
            let worksheet = workbook.add_worksheet();

            let result = write_cell(worksheet, 0, 0, &Value::Blob(blob.clone()), BlobHandling::Placeholder);
            assert!(result.is_ok(), "Failed for blob size {}: {}", blob.len(), result.unwrap_err());
        }
    }
}
