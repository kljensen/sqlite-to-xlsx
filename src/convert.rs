use anyhow::{Context, Result};
use rust_xlsxwriter::{Format, Workbook, Worksheet};
use rusqlite::{Connection, types::Value, OpenFlags};
use std::{collections::HashSet, fmt::Write, io::IsTerminal, path::Path, time::Instant};

use crate::{discover_tables, sanitize_sheet_name, TableInfo};

/// Configuration for how BLOB values should be written to Excel cells
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BlobHandling {
    /// Write a placeholder string like "[BLOB: 123 bytes]"
    #[default]
    Placeholder,
    /// Write hexadecimal representation like "0x48656c6c6f"
    Hex,
    /// Write base64 encoded string
    Base64,
    /// Write an empty cell (skip the BLOB value)
    Skip,
}

/// Maximum safe integer for f64 (2^53)
/// Numbers larger than this cannot be precisely represented in f64
const MAX_SAFE_INTEGER: i64 = 9_007_199_254_740_992; // 2^53

/// Maximum string length before truncation (Excel has a 32,767 char limit)
const MAX_STRING_LENGTH: usize = 32_767;

/// Truncation suffix for overly long strings
const TRUNCATION_SUFFIX: &str = "... [truncated]";

/// Minimum row count for showing a progress bar during table export
const PROGRESS_BAR_THRESHOLD: u64 = 1000;

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
    if !(-MAX_SAFE_INTEGER..=MAX_SAFE_INTEGER).contains(&value) {
        eprintln!(
            "Warning: Large integer {} at row {}, col {} exceeds f64 precision, writing as string",
            value, row, col
        );
        worksheet
            .write_string(row, col, value.to_string())
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
/// Sanitizes strings that contain binary data or control characters,
/// then truncates if they exceed Excel's maximum length (32,767 characters)
/// and logs a warning.
fn write_string(worksheet: &mut Worksheet, row: u32, col: u16, value: &str) -> Result<()> {
    // Sanitize the string first to handle binary data stored as text
    let sanitized = sanitize_string(value);
    let char_count = sanitized.chars().count();

    if char_count > MAX_STRING_LENGTH {
        eprintln!(
            "Warning: String at row {}, col {} exceeds Excel max length ({} chars), truncating",
            row, col, char_count
        );
        let truncated = truncate_string(&sanitized, MAX_STRING_LENGTH);
        worksheet
            .write_string(row, col, &truncated)
            .map_err(|e| anyhow::anyhow!(e))?;
    } else {
        worksheet
            .write_string(row, col, &sanitized)
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

/// Sanitizes a string for Excel compatibility
///
/// Handles binary data stored as text by:
/// - Replacing the Unicode replacement character (U+FFFD) with a placeholder
/// - Removing null bytes and other control characters (except tab, newline, carriage return)
/// - Detecting binary content (like PDF data) and replacing with a placeholder
fn sanitize_string(s: &str) -> String {
    // Check for common binary file signatures stored as text
    if s.starts_with("%PDF") || s.starts_with("\x00") || s.starts_with("PK\x03\x04") {
        return format!("[Binary data: {} bytes]", s.len());
    }

    // Count replacement characters - high count indicates binary data
    let replacement_count = s.chars().filter(|&c| c == '\u{FFFD}').count();
    if replacement_count > s.chars().count() / 10 {
        // More than 10% replacement chars suggests binary data
        return format!("[Binary data: {} bytes]", s.len());
    }

    // Sanitize: keep printable chars, tabs, newlines, carriage returns
    let sanitized: String = s
        .chars()
        .map(|c| {
            if c == '\u{FFFD}' {
                '?' // Replace Unicode replacement char
            } else if c == '\t' || c == '\n' || c == '\r' {
                c // Keep whitespace
            } else if c.is_control() {
                ' ' // Replace other control chars with space
            } else {
                c
            }
        })
        .collect();

    sanitized
}

/// Truncates a string to fit within the maximum length
///
/// Preserves room for the truncation suffix indicator.
fn truncate_string(s: &str, max_len: usize) -> String {
    // If string fits within max_len, return it as-is
    let char_count = s.chars().count();
    if char_count <= max_len {
        return s.to_string();
    }
    // If max_len is too small for the suffix, return just the suffix
    let suffix_len = TRUNCATION_SUFFIX.chars().count();
    if max_len <= suffix_len {
        return TRUNCATION_SUFFIX.to_string();
    }
    // Truncate the string and add suffix
    let available = max_len - suffix_len;
    let mut result = String::with_capacity(s.len().min(max_len));
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

/// Calculates the display width (character count) of a SQLite value
fn value_display_width(value: &Value, blob_handling: BlobHandling) -> usize {
    match value {
        Value::Null => 0,
        Value::Integer(i) => i.to_string().chars().count(),
        Value::Real(f) => {
            if f.is_nan() {
                3 // "NaN"
            } else if f.is_infinite() {
                if f.is_sign_positive() { 8 } else { 9 } // "Infinity" or "-Infinity"
            } else {
                f.to_string().chars().count()
            }
        }
        Value::Text(s) => s.chars().count(),
        Value::Blob(b) => {
            match blob_handling {
                BlobHandling::Placeholder => format!("[BLOB: {} bytes]", b.len()).chars().count(),
                BlobHandling::Hex => b.len() * 2 + 2, // "0x" + 2 chars per byte
                BlobHandling::Base64 => b.len().div_ceil(3) * 4, // Base64 formula: ceil(n/3) * 4
                BlobHandling::Skip => 0,
            }
        }
    }
}

/// Configuration options for the conversion process
#[derive(Debug, Clone)]
pub struct ConvertOptions {
    /// Include only these tables (None means all tables)
    pub tables: Option<Vec<String>>,
    /// Exclude these tables (None means no exclusions)
    pub exclude: Option<Vec<String>>,
    /// Exclude specific columns (None means no exclusions)
    /// Format: "column" for all tables, "table.column" for specific table
    pub exclude_columns: Option<Vec<String>>,
    /// How to handle BLOB values
    pub blob_handling: BlobHandling,
    /// Whether to write column headers
    pub write_headers: bool,
    /// Suppress all output except errors
    pub quiet: bool,
    /// Custom SQL queries with their target sheet names (query, sheet_name)
    pub queries: Vec<(String, String)>,
}

impl Default for ConvertOptions {
    fn default() -> Self {
        ConvertOptions {
            tables: None,
            exclude: None,
            exclude_columns: None,
            blob_handling: BlobHandling::default(),
            write_headers: true,
            quiet: false,
            queries: Vec::new(),
        }
    }
}

/// Checks if a column should be excluded based on the exclusion list
///
/// Supports two formats:
/// - "column_name" - matches column in any table (case-insensitive)
/// - "table.column" - matches column only in specific table (case-insensitive)
fn should_exclude_column(table_name: &str, column_name: &str, exclusions: &Option<Vec<String>>) -> bool {
    let Some(ref excludes) = exclusions else {
        return false;
    };

    let table_lower = table_name.to_lowercase();
    let column_lower = column_name.to_lowercase();

    for pattern in excludes {
        let pattern_lower = pattern.to_lowercase();
        if let Some((tbl, col)) = pattern_lower.split_once('.') {
            // table.column format - match specific table
            if tbl == table_lower && col == column_lower {
                return true;
            }
        } else {
            // column only - match in any table
            if pattern_lower == column_lower {
                return true;
            }
        }
    }
    false
}

/// Statistics about the conversion process
#[derive(Debug, Clone)]
pub struct ConvertStats {
    /// Number of tables successfully exported
    pub tables_exported: usize,
    /// Total number of rows written across all tables
    pub total_rows: usize,
    /// Time taken for the conversion
    pub duration: std::time::Duration,
}

/// Converts SQLite database to XLSX format
///
/// # Arguments
/// * `input_path` - Path to the input SQLite database file
/// * `output_path` - Path where the XLSX file will be written
/// * `options` - Configuration options for the conversion
///
/// # Returns
/// Statistics about the conversion process
///
/// # Examples
/// ```no_run
/// use sqlite_to_xlsx::{convert, ConvertOptions};
/// use std::path::Path;
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let options = ConvertOptions::default();
/// let stats = convert(
///     Path::new("database.db"),
///     Path::new("output.xlsx"),
///     &options
/// )?;
/// println!("Exported {} tables with {} rows", stats.tables_exported, stats.total_rows);
/// # Ok(())
/// # }
/// ```
#[must_use = "conversion stats should be used or the result should be checked for errors"]
pub fn convert(input_path: &Path, output_path: &Path, options: &ConvertOptions) -> Result<ConvertStats> {
    let start = Instant::now();

    // 1. Open SQLite database (read-only)
    let flags = OpenFlags::SQLITE_OPEN_READ_ONLY | OpenFlags::SQLITE_OPEN_NO_MUTEX;
    let conn = Connection::open_with_flags(input_path, flags)
        .with_context(|| format!("Cannot open database: '{}'", input_path.display()))?;

    // 2. Create new Workbook
    let mut workbook = Workbook::new();

    // 3. Discover all tables using discover_tables()
    let mut tables = discover_tables(&conn)?;
    tables.sort_by(|a, b| a.name.cmp(&b.name));

    // Filter tables based on options
    let tables: Vec<TableInfo> = if let Some(ref include) = options.tables {
        tables.into_iter()
            .filter(|t| include.contains(&t.name))
            .collect()
    } else {
        tables
    };

    let tables: Vec<TableInfo> = if let Some(ref exclude) = options.exclude {
        tables.into_iter()
            .filter(|t| !exclude.contains(&t.name))
            .collect()
    } else {
        tables
    };

    let mut exported_count = 0;
    let mut total_rows = 0;
    let mut used_sheet_names = HashSet::new();

    // Determine if we should show progress bars
    let is_tty = std::io::stdout().is_terminal();
    let show_progress = is_tty && !options.quiet;

    // 4. For each table (sorted alphabetically)
    for table_info in tables {
        // Filter columns based on exclusion list
        let columns: Vec<_> = table_info.columns.iter()
            .filter(|col| !should_exclude_column(&table_info.name, &col.name, &options.exclude_columns))
            .collect();

        // Skip table if all columns are excluded
        if columns.is_empty() {
            if !options.quiet {
                eprintln!("Warning: All columns excluded from table '{}', skipping", table_info.name);
            }
            continue;
        }

        // Get row count for progress decision
        let sql_count = format!("SELECT COUNT(*) FROM [{}]", table_info.name.replace("'", "''"));
        let row_count_estimate: i64 = conn.query_row(&sql_count, [], |row| row.get(0))
            .unwrap_or(0);
        let is_large_table = row_count_estimate > PROGRESS_BAR_THRESHOLD as i64;

        // a. Sanitize name using sanitize_sheet_name()
        let sheet_name = sanitize_sheet_name(&table_info.name, &mut used_sheet_names);

        // b. Create worksheet
        let worksheet = workbook.add_worksheet().set_name(&sheet_name)
            .map_err(|e| anyhow::anyhow!("Failed to create sheet '{}': {}", sheet_name, e))?;

        // Track column widths for auto-fit
        let num_columns = columns.len();
        let mut column_widths: Vec<usize> = vec![0; num_columns];

        // c. Write column headers (row 0) with bold format
        if options.write_headers {
            let bold_format = Format::new().set_bold();
            for (col, column) in columns.iter().enumerate() {
                worksheet.write_string_with_format(0, col as u16, &column.name, &bold_format)
                    .map_err(|e| anyhow::anyhow!("Failed to write header for '{}': {}", column.name, e))?;
                // Track header width
                column_widths[col] = column.name.chars().count();
            }
        }

        // d. Build SELECT with specific columns (to handle exclusions)
        // e. Stream rows using write_cell()
        let column_list: String = columns.iter()
            .map(|c| format!("[{}]", c.name.replace("]", "]]")))
            .collect::<Vec<_>>()
            .join(", ");
        let sql = format!("SELECT {} FROM [{}]", column_list, table_info.name.replace("'", "''"));
        let mut stmt = conn.prepare(&sql)?;

        let mut row_count: usize = 0;
        let header_offset = if options.write_headers { 1 } else { 0 };

        // Create progress bar for large tables if TTY and not quiet
        let progress = if show_progress && is_large_table {
            use indicatif::ProgressBar;
            let pb = ProgressBar::new(row_count_estimate as u64);
            pb.set_style(indicatif::ProgressStyle::default_bar()
                .template("[{elapsed_precise}] [{bar:20.cyan/blue}] {pos}/{len} ({percent}%)")
                .map_err(|e| anyhow::anyhow!("Failed to set progress style: {}", e))?
                .progress_chars("=>-"));
            pb.set_message(format!("Exporting {}", table_info.name));
            Some(pb)
        } else {
            None
        };

        let mut rows = stmt.query([])?;
        while let Some(row) = rows.next()? {
            let row_idx = (row_count + header_offset) as u32;

            for (col_idx, col_width) in column_widths.iter_mut().enumerate() {
                let value: Value = row.get(col_idx)?;
                write_cell(worksheet, row_idx, col_idx as u16, &value, options.blob_handling)
                    .map_err(|e| anyhow::anyhow!("Failed to write cell at row {}, col {}: {}", row_idx, col_idx, e))?;
                // Update column width
                let width = value_display_width(&value, options.blob_handling);
                *col_width = (*col_width).max(width);
            }
            row_count += 1;

            // Update progress bar
            if let Some(ref pb) = progress {
                pb.inc(1);
            }
        }

        // Set column widths (auto-fit)
        for (col, &width) in column_widths.iter().enumerate() {
            // Formula: min(50, max(8, chars * 1.1))
            let col_width = (50.0_f64).min((8.0_f64).max(width as f64 * 1.1));
            worksheet.set_column_width(col as u16, col_width)
                .map_err(|e| anyhow::anyhow!("Failed to set column width for column {}: {}", col, e))?;
        }

        // Finish progress bar
        if let Some(pb) = progress {
            pb.finish();
        }

        // Print status for small tables if not quiet
        if !options.quiet && !is_large_table {
            println!("Exported: {} ({} rows)", table_info.name, row_count);
        }

        total_rows += row_count;
        exported_count += 1;
    }

    // 5. Execute custom queries and export to named sheets
    for (query, sheet_name) in &options.queries {
        // Sanitize sheet name
        let sanitized_name = sanitize_sheet_name(sheet_name, &mut used_sheet_names);

        // Create worksheet
        let worksheet = workbook.add_worksheet().set_name(&sanitized_name)
            .map_err(|e| anyhow::anyhow!("Failed to create sheet '{}': {}", sanitized_name, e))?;

        // Execute query and write to sheet, get column widths
        let (query_rows, column_widths) = query_to_sheet(&conn, query, &sanitized_name, worksheet, options)
            .with_context(|| format!("Failed to execute query for sheet '{}': {}", sanitized_name, query))?;

        // Set column widths (auto-fit)
        for (col, &width) in column_widths.iter().enumerate() {
            let col_width = (50.0_f64).min((8.0_f64).max(width as f64 * 1.1));
            worksheet.set_column_width(col as u16, col_width)
                .map_err(|e| anyhow::anyhow!("Failed to set column width for column {}: {}", col, e))?;
        }

        total_rows += query_rows;

        // Print status if not quiet
        if !options.quiet {
            println!("Exported: {} ({} rows from query)", sanitized_name, query_rows);
        }
    }

    // 6. Save workbook
    workbook.save(output_path)
        .with_context(|| format!("Cannot write to output file: '{}'", output_path.display()))?;

    // Print final summary if not quiet
    if !options.quiet {
        println!("✓ Exported {} tables ({} rows) to {} in {:.1}s",
            exported_count,
            total_rows,
            output_path.display(),
            start.elapsed().as_secs_f64()
        );
    }

    Ok(ConvertStats {
        tables_exported: exported_count,
        total_rows,
        duration: start.elapsed(),
    })
}

/// Executes a custom SQL query and exports results to a named worksheet
///
/// # Arguments
/// * `conn` - SQLite connection
/// * `query` - SQL query string to execute
/// * `sheet_name` - Name for the worksheet
/// * `worksheet` - The Excel worksheet to write to
/// * `options` - Conversion options
///
/// # Returns
/// A tuple of (number of rows written, column widths for auto-fit)
///
/// # Errors
/// Returns an error if:
/// - The SQL query is invalid
/// - The query execution fails
/// - Writing to the worksheet fails
pub fn query_to_sheet(
    conn: &Connection,
    query: &str,
    sheet_name: &str,
    worksheet: &mut Worksheet,
    options: &ConvertOptions,
) -> Result<(usize, Vec<usize>)> {
    // Prepare and execute the query
    let mut stmt = conn.prepare(query)
        .with_context(|| format!("Query failed for sheet '{}': {}", sheet_name, query))?;

    let column_count = stmt.column_count();
    let mut row_count: usize = 0;

    // Track column widths for auto-fit
    let mut column_widths: Vec<usize> = vec![0; column_count];

    // Write column headers if requested
    if options.write_headers {
        let bold_format = Format::new().set_bold();
        for (col, col_width) in column_widths.iter_mut().enumerate() {
            let column_name = stmt.column_name(col)
                .unwrap_or("Column");
            worksheet.write_string_with_format(0, col as u16, column_name, &bold_format)
                .map_err(|e| anyhow::anyhow!("Failed to write header for column {}: {}", col, e))?;
            // Track header width
            *col_width = column_name.chars().count();
        }
    }

    let header_offset = if options.write_headers { 1 } else { 0 };

    // Stream rows using write_cell()
    let mut rows = stmt.query([])?;
    while let Some(row) = rows.next()? {
        let row_idx = (row_count + header_offset) as u32;

        for (col_idx, col_width) in column_widths.iter_mut().enumerate() {
            let value: Value = row.get(col_idx)?;
            write_cell(worksheet, row_idx, col_idx as u16, &value, options.blob_handling)
                .map_err(|e| anyhow::anyhow!("Failed to write cell at row {}, col {}: {}", row_idx, col_idx, e))?;
            // Update column width
            let width = value_display_width(&value, options.blob_handling);
            *col_width = (*col_width).max(width);
        }
        row_count += 1;
    }

    Ok((row_count, column_widths))
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

        let result = write_cell(worksheet, 0, 0, &Value::Real(1.5), BlobHandling::default());
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
    fn test_sanitize_string_normal_text() {
        // Normal text should pass through unchanged
        let normal = "Hello, World!";
        assert_eq!(sanitize_string(normal), normal);

        // Text with tabs and newlines should be preserved
        let with_whitespace = "Line 1\tcolumn\nLine 2";
        assert_eq!(sanitize_string(with_whitespace), with_whitespace);
    }

    #[test]
    fn test_sanitize_string_pdf_detection() {
        // PDF-like content should be replaced with placeholder
        let pdf = "%PDF-1.7 some binary content here";
        let sanitized = sanitize_string(pdf);
        assert!(sanitized.starts_with("[Binary data:"));
        assert!(sanitized.ends_with("bytes]"));
    }

    #[test]
    fn test_sanitize_string_null_byte_detection() {
        // Content starting with null byte should be detected as binary
        let with_null = "\x00Some binary data";
        let sanitized = sanitize_string(with_null);
        assert!(sanitized.starts_with("[Binary data:"));
    }

    #[test]
    fn test_sanitize_string_control_chars() {
        // Control characters (except \t, \n, \r) should be replaced with space
        let with_control = "Hello\x01World\x02Test";
        let sanitized = sanitize_string(with_control);
        assert_eq!(sanitized, "Hello World Test");
    }

    #[test]
    fn test_sanitize_string_replacement_char() {
        // Unicode replacement character should be replaced with ?
        let with_replacement = "Hello\u{FFFD}World";
        let sanitized = sanitize_string(with_replacement);
        assert_eq!(sanitized, "Hello?World");
    }

    #[test]
    fn test_sanitize_string_high_replacement_ratio() {
        // String with >10% replacement chars should be treated as binary
        let mostly_invalid = "\u{FFFD}\u{FFFD}ab\u{FFFD}";
        let sanitized = sanitize_string(mostly_invalid);
        assert!(sanitized.starts_with("[Binary data:"));
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
    fn test_should_exclude_column_none() {
        // No exclusions - should never exclude
        assert!(!should_exclude_column("users", "email", &None));
        assert!(!should_exclude_column("products", "price", &None));
    }

    #[test]
    fn test_should_exclude_column_by_name() {
        // Column name only - matches any table
        let excludes = Some(vec!["password".to_string()]);
        assert!(should_exclude_column("users", "password", &excludes));
        assert!(should_exclude_column("admins", "password", &excludes));
        assert!(!should_exclude_column("users", "email", &excludes));
    }

    #[test]
    fn test_should_exclude_column_by_table_column() {
        // table.column - matches only specific table
        let excludes = Some(vec!["users.password".to_string()]);
        assert!(should_exclude_column("users", "password", &excludes));
        assert!(!should_exclude_column("admins", "password", &excludes));
        assert!(!should_exclude_column("users", "email", &excludes));
    }

    #[test]
    fn test_should_exclude_column_case_insensitive() {
        // Should be case-insensitive
        let excludes = Some(vec!["PASSWORD".to_string(), "Users.Email".to_string()]);
        assert!(should_exclude_column("users", "password", &excludes));
        assert!(should_exclude_column("USERS", "PASSWORD", &excludes));
        assert!(should_exclude_column("users", "email", &excludes));
        assert!(should_exclude_column("Users", "EMAIL", &excludes));
    }

    #[test]
    fn test_should_exclude_column_multiple() {
        // Multiple exclusions
        let excludes = Some(vec![
            "password".to_string(),
            "users.secret".to_string(),
            "products.internal_code".to_string(),
        ]);
        assert!(should_exclude_column("users", "password", &excludes));
        assert!(should_exclude_column("products", "password", &excludes));
        assert!(should_exclude_column("users", "secret", &excludes));
        assert!(!should_exclude_column("products", "secret", &excludes));
        assert!(should_exclude_column("products", "internal_code", &excludes));
        assert!(!should_exclude_column("users", "internal_code", &excludes));
    }

    #[test]
    fn test_multiple_cells_same_row() {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let values = [
            Value::Null,
            Value::Integer(42),
            Value::Real(1.5),
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
        let worksheet = workbook.add_worksheet();

        let values = [
            Value::Integer(1),
            Value::Integer(2),
            Value::Integer(3),
        ];

        for (row, value) in values.iter().enumerate() {
            let result = write_cell(worksheet, row as u32, 0, value, BlobHandling::default());
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
        let worksheet = workbook.add_worksheet();

        // Test positive and negative zero (should both work)
        let result = write_cell(worksheet, 0, 0, &Value::Real(0.0), BlobHandling::default());
        assert!(result.is_ok());

        let result = write_cell(worksheet, 0, 1, &Value::Real(-0.0), BlobHandling::default());
        assert!(result.is_ok());

        // Test subnormal numbers
        let result = write_cell(worksheet, 0, 2, &Value::Real(1e-320), BlobHandling::default());
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

    #[test]
    fn test_convert_options_default() {
        let options = ConvertOptions::default();
        assert!(options.tables.is_none());
        assert!(options.exclude.is_none());
        assert_eq!(options.blob_handling, BlobHandling::Placeholder);
        assert!(options.write_headers);
    }

    #[test]
    fn test_convert_integration() {
        use tempfile::TempDir;
        use rusqlite::Connection;

        // Create a temporary directory
        let temp_dir = TempDir::new().expect("Failed to create temp dir");
        let db_path = temp_dir.path().join("test.db");
        let xlsx_path = temp_dir.path().join("test.xlsx");

        // Create a test SQLite database with multiple tables
        let conn = Connection::open(&db_path).expect("Failed to create database");

        // Create users table
        conn.execute(
            "CREATE TABLE users (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                email TEXT,
                age INTEGER,
                balance REAL
            )",
            [],
        ).expect("Failed to create users table");

        conn.execute(
            "INSERT INTO users (id, name, email, age, balance) VALUES
                (1, 'Alice', 'alice@example.com', 30, 100.50),
                (2, 'Bob', 'bob@example.com', 25, 75.25),
                (3, 'Charlie', NULL, 35, 0.0)",
            [],
        ).expect("Failed to insert users");

        // Create products table
        conn.execute(
            "CREATE TABLE products (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                price REAL
            )",
            [],
        ).expect("Failed to create products table");

        conn.execute(
            "INSERT INTO products (id, name, price) VALUES
                (1, 'Widget', 19.99),
                (2, 'Gadget', 29.99)",
            [],
        ).expect("Failed to insert products");

        // Drop the connection so the file is fully written
        drop(conn);

        // Convert the database
        let options = ConvertOptions::default();
        let stats = convert(&db_path, &xlsx_path, &options)
            .expect("Failed to convert database");

        // Verify stats
        assert_eq!(stats.tables_exported, 2);
        assert_eq!(stats.total_rows, 5); // 3 users + 2 products

        // Verify output file exists
        assert!(xlsx_path.exists(), "XLSX file should exist");
    }

    #[test]
    fn test_convert_with_table_filter() {
        use tempfile::TempDir;
        use rusqlite::Connection;

        let temp_dir = TempDir::new().expect("Failed to create temp dir");
        let db_path = temp_dir.path().join("test.db");
        let xlsx_path = temp_dir.path().join("test.xlsx");

        let conn = Connection::open(&db_path).expect("Failed to create database");

        conn.execute(
            "CREATE TABLE table_a (id INTEGER, value TEXT)",
            [],
        ).expect("Failed to create table_a");

        conn.execute(
            "CREATE TABLE table_b (id INTEGER, value TEXT)",
            [],
        ).expect("Failed to create table_b");

        conn.execute(
            "CREATE TABLE table_c (id INTEGER, value TEXT)",
            [],
        ).expect("Failed to create table_c");

        conn.execute("INSERT INTO table_a VALUES (1, 'a')", []).expect("Failed to insert");
        conn.execute("INSERT INTO table_b VALUES (2, 'b')", []).expect("Failed to insert");
        conn.execute("INSERT INTO table_c VALUES (3, 'c')", []).expect("Failed to insert");

        drop(conn);

        // Only convert table_a and table_b
        let options = ConvertOptions {
            tables: Some(vec!["table_a".to_string(), "table_b".to_string()]),
            ..Default::default()
        };

        let stats = convert(&db_path, &xlsx_path, &options)
            .expect("Failed to convert database");

        assert_eq!(stats.tables_exported, 2);
        assert_eq!(stats.total_rows, 2);
        assert!(xlsx_path.exists());
    }

    #[test]
    fn test_convert_with_exclude() {
        use tempfile::TempDir;
        use rusqlite::Connection;

        let temp_dir = TempDir::new().expect("Failed to create temp dir");
        let db_path = temp_dir.path().join("test.db");
        let xlsx_path = temp_dir.path().join("test.xlsx");

        let conn = Connection::open(&db_path).expect("Failed to create database");

        conn.execute(
            "CREATE TABLE keep_this (id INTEGER)",
            [],
        ).expect("Failed to create table");

        conn.execute(
            "CREATE TABLE skip_this (id INTEGER)",
            [],
        ).expect("Failed to create table");

        conn.execute("INSERT INTO keep_this VALUES (1)", []).expect("Failed to insert");
        conn.execute("INSERT INTO skip_this VALUES (2)", []).expect("Failed to insert");

        drop(conn);

        let options = ConvertOptions {
            exclude: Some(vec!["skip_this".to_string()]),
            ..Default::default()
        };

        let stats = convert(&db_path, &xlsx_path, &options)
            .expect("Failed to convert database");

        assert_eq!(stats.tables_exported, 1);
        assert_eq!(stats.total_rows, 1);
    }

    #[test]
    fn test_convert_without_headers() {
        use tempfile::TempDir;
        use rusqlite::Connection;

        let temp_dir = TempDir::new().expect("Failed to create temp dir");
        let db_path = temp_dir.path().join("test.db");
        let xlsx_path = temp_dir.path().join("test.xlsx");

        let conn = Connection::open(&db_path).expect("Failed to create database");
        conn.execute("CREATE TABLE test (id INTEGER, name TEXT)", []).expect("Failed to create table");
        conn.execute("INSERT INTO test VALUES (1, 'Alice')", []).expect("Failed to insert");

        drop(conn);

        let options = ConvertOptions {
            write_headers: false,
            ..Default::default()
        };

        let stats = convert(&db_path, &xlsx_path, &options)
            .expect("Failed to convert database");

        assert_eq!(stats.tables_exported, 1);
        assert_eq!(stats.total_rows, 1);
    }

    #[test]
    fn test_convert_blob_handling() {
        use tempfile::TempDir;
        use rusqlite::Connection;

        let temp_dir = TempDir::new().expect("Failed to create temp dir");
        let db_path = temp_dir.path().join("test.db");
        let xlsx_path = temp_dir.path().join("test.xlsx");

        let conn = Connection::open(&db_path).expect("Failed to create database");
        conn.execute("CREATE TABLE blob_test (id INTEGER, data BLOB)", []).expect("Failed to create table");
        conn.execute("INSERT INTO blob_test VALUES (1, x'48656c6c6f')", []).expect("Failed to insert");

        drop(conn);

        // Test with Skip mode
        let options = ConvertOptions {
            blob_handling: BlobHandling::Skip,
            ..Default::default()
        };

        let stats = convert(&db_path, &xlsx_path, &options)
            .expect("Failed to convert database");

        assert_eq!(stats.tables_exported, 1);
        assert_eq!(stats.total_rows, 1);
    }

    #[test]
    fn test_convert_empty_database() {
        use tempfile::TempDir;
        use rusqlite::Connection;

        let temp_dir = TempDir::new().expect("Failed to create temp dir");
        let db_path = temp_dir.path().join("test.db");
        let xlsx_path = temp_dir.path().join("test.xlsx");

        let conn = Connection::open(&db_path).expect("Failed to create database");
        drop(conn);

        let options = ConvertOptions::default();
        let stats = convert(&db_path, &xlsx_path, &options)
            .expect("Failed to convert empty database");

        assert_eq!(stats.tables_exported, 0);
        assert_eq!(stats.total_rows, 0);
        assert!(xlsx_path.exists());
    }

    #[test]
    fn test_convert_with_column_exclusion() {
        use tempfile::TempDir;
        use rusqlite::Connection;
        use calamine::{Reader, open_workbook, Xlsx, DataType};

        let temp_dir = TempDir::new().expect("Failed to create temp dir");
        let db_path = temp_dir.path().join("test.db");
        let xlsx_path = temp_dir.path().join("test.xlsx");

        let conn = Connection::open(&db_path).expect("Failed to create database");

        conn.execute(
            "CREATE TABLE users (id INTEGER, name TEXT, password TEXT, email TEXT)",
            [],
        ).expect("Failed to create users table");

        conn.execute(
            "CREATE TABLE admins (id INTEGER, name TEXT, password TEXT)",
            [],
        ).expect("Failed to create admins table");

        conn.execute("INSERT INTO users VALUES (1, 'Alice', 'secret123', 'alice@example.com')", [])
            .expect("Failed to insert");
        conn.execute("INSERT INTO admins VALUES (1, 'Admin', 'admin456')", [])
            .expect("Failed to insert");

        drop(conn);

        // Exclude 'password' from all tables
        let options = ConvertOptions {
            exclude_columns: Some(vec!["password".to_string()]),
            ..Default::default()
        };

        let stats = convert(&db_path, &xlsx_path, &options)
            .expect("Failed to convert database");

        assert_eq!(stats.tables_exported, 2);

        // Verify the exported Excel doesn't have the password column
        let mut workbook: Xlsx<_> = open_workbook(&xlsx_path).expect("Failed to open xlsx");

        // Check users sheet - should have id, name, email (no password)
        let users_range = workbook.worksheet_range("users").expect("Failed to get users sheet");
        let headers: Vec<String> = users_range.rows().next().unwrap()
            .iter().filter_map(|c| c.get_string().map(String::from)).collect();
        assert_eq!(headers, vec!["id", "name", "email"]);
        assert!(!headers.contains(&"password".to_string()));

        // Check admins sheet - should have id, name (no password)
        let admins_range = workbook.worksheet_range("admins").expect("Failed to get admins sheet");
        let headers: Vec<String> = admins_range.rows().next().unwrap()
            .iter().filter_map(|c| c.get_string().map(String::from)).collect();
        assert_eq!(headers, vec!["id", "name"]);
        assert!(!headers.contains(&"password".to_string()));
    }

    #[test]
    fn test_convert_with_table_specific_column_exclusion() {
        use tempfile::TempDir;
        use rusqlite::Connection;
        use calamine::{Reader, open_workbook, Xlsx, DataType};

        let temp_dir = TempDir::new().expect("Failed to create temp dir");
        let db_path = temp_dir.path().join("test.db");
        let xlsx_path = temp_dir.path().join("test.xlsx");

        let conn = Connection::open(&db_path).expect("Failed to create database");

        conn.execute(
            "CREATE TABLE users (id INTEGER, secret TEXT)",
            [],
        ).expect("Failed to create users table");

        conn.execute(
            "CREATE TABLE products (id INTEGER, secret TEXT)",
            [],
        ).expect("Failed to create products table");

        conn.execute("INSERT INTO users VALUES (1, 'user_secret')", []).expect("Failed to insert");
        conn.execute("INSERT INTO products VALUES (1, 'product_secret')", []).expect("Failed to insert");

        drop(conn);

        // Exclude 'secret' only from users table
        let options = ConvertOptions {
            exclude_columns: Some(vec!["users.secret".to_string()]),
            ..Default::default()
        };

        let stats = convert(&db_path, &xlsx_path, &options)
            .expect("Failed to convert database");

        assert_eq!(stats.tables_exported, 2);

        let mut workbook: Xlsx<_> = open_workbook(&xlsx_path).expect("Failed to open xlsx");

        // Users should not have 'secret'
        let users_range = workbook.worksheet_range("users").expect("Failed to get users sheet");
        let headers: Vec<String> = users_range.rows().next().unwrap()
            .iter().filter_map(|c| c.get_string().map(String::from)).collect();
        assert_eq!(headers, vec!["id"]);

        // Products should still have 'secret'
        let products_range = workbook.worksheet_range("products").expect("Failed to get products sheet");
        let headers: Vec<String> = products_range.rows().next().unwrap()
            .iter().filter_map(|c| c.get_string().map(String::from)).collect();
        assert_eq!(headers, vec!["id", "secret"]);
    }
}
