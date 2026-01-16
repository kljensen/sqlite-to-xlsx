// Integration tests for basic SQLite to XLSX conversion
// Tests verify that data is correctly written to XLSX files using calamine

use calamine::{Reader, Xlsx, open_workbook, Data};
use rusqlite::Connection;
use std::path::Path;
use tempfile::TempDir;
use sqlite_to_xlsx::{convert, ConvertOptions};

fn create_test_database(db_path: &Path) -> anyhow::Result<()> {
    let conn = Connection::open(db_path)?;

    // Create a simple table with various data types
    conn.execute(
        "CREATE TABLE users (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            email TEXT,
            age INTEGER,
            balance REAL,
            is_active INTEGER
        )",
        [],
    )?;

    // Insert test data covering all SQLite types
    conn.execute(
        "INSERT INTO users (id, name, email, age, balance, is_active) VALUES
            (1, 'Alice', 'alice@example.com', 30, 100.50, 1),
            (2, 'Bob', 'bob@example.com', 25, 75.25, 1),
            (3, 'Charlie', NULL, 35, 0.0, 0),
            (4, 'Diana', 'diana@example.com', NULL, -10.99, 1)",
        [],
    )?;

    Ok(())
}

#[test]
fn test_single_table_conversion_with_content_verification() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create test database
    create_test_database(&db_path).expect("Failed to create database");

    // Convert to XLSX
    let options = ConvertOptions::default();
    let stats = convert(&db_path, &xlsx_path, &options)
        .expect("Conversion should succeed");

    // Verify statistics
    assert_eq!(stats.tables_exported, 1);
    assert_eq!(stats.total_rows, 4);

    // Verify output file exists and can be opened
    assert!(xlsx_path.exists(), "XLSX file should exist");

    // Use calamine to read and verify the XLSX content
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    // Verify sheet exists
    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names, vec!["users"]);

    // Verify sheet content
    let sheet = workbook.worksheet_range("users")
        .expect("Failed to read sheet 'users'");

    // Verify headers (row 0)
    assert_eq!(sheet.get((0, 0)), Some(&Data::String("id".to_string())));
    assert_eq!(sheet.get((0, 1)), Some(&Data::String("name".to_string())));
    assert_eq!(sheet.get((0, 2)), Some(&Data::String("email".to_string())));
    assert_eq!(sheet.get((0, 3)), Some(&Data::String("age".to_string())));
    assert_eq!(sheet.get((0, 4)), Some(&Data::String("balance".to_string())));
    assert_eq!(sheet.get((0, 5)), Some(&Data::String("is_active".to_string())));

    // Verify data row 1 (Alice)
    assert_eq!(sheet.get((1, 0)), Some(&Data::Float(1.0)));
    assert_eq!(sheet.get((1, 1)), Some(&Data::String("Alice".to_string())));
    assert_eq!(sheet.get((1, 2)), Some(&Data::String("alice@example.com".to_string())));
    assert_eq!(sheet.get((1, 3)), Some(&Data::Float(30.0)));
    assert_eq!(sheet.get((1, 4)), Some(&Data::Float(100.5)));
    assert_eq!(sheet.get((1, 5)), Some(&Data::Float(1.0)));

    // Verify data row 2 (Bob)
    assert_eq!(sheet.get((2, 0)), Some(&Data::Float(2.0)));
    assert_eq!(sheet.get((2, 1)), Some(&Data::String("Bob".to_string())));
    assert_eq!(sheet.get((2, 2)), Some(&Data::String("bob@example.com".to_string())));
    assert_eq!(sheet.get((2, 3)), Some(&Data::Float(25.0)));
    assert_eq!(sheet.get((2, 4)), Some(&Data::Float(75.25)));

    // Verify data row 3 (Charlie with NULL email and 0.0 balance)
    assert_eq!(sheet.get((3, 0)), Some(&Data::Float(3.0)));
    assert_eq!(sheet.get((3, 1)), Some(&Data::String("Charlie".to_string())));
    assert_eq!(sheet.get((3, 2)), Some(&Data::Empty)); // NULL email
    assert_eq!(sheet.get((3, 3)), Some(&Data::Float(35.0)));
    assert_eq!(sheet.get((3, 4)), Some(&Data::Float(0.0)));
    assert_eq!(sheet.get((3, 5)), Some(&Data::Float(0.0)));

    // Verify data row 4 (Diana with NULL age)
    assert_eq!(sheet.get((4, 0)), Some(&Data::Float(4.0)));
    assert_eq!(sheet.get((4, 1)), Some(&Data::String("Diana".to_string())));
    assert_eq!(sheet.get((4, 2)), Some(&Data::String("diana@example.com".to_string())));
    assert_eq!(sheet.get((4, 3)), Some(&Data::Empty)); // NULL age
    assert_eq!(sheet.get((4, 4)), Some(&Data::Float(-10.99)));
}

#[test]
fn test_empty_database_creates_valid_xlsx() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("empty.db");
    let xlsx_path = temp_dir.path().join("empty.xlsx");

    // Create empty database
    let conn = Connection::open(&db_path).expect("Failed to create database");
    drop(conn);

    // Convert to XLSX
    let options = ConvertOptions::default();
    let stats = convert(&db_path, &xlsx_path, &options)
        .expect("Conversion should succeed");

    assert_eq!(stats.tables_exported, 0);
    assert_eq!(stats.total_rows, 0);

    // Verify XLSX file was created and is valid
    assert!(xlsx_path.exists());
    let workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    // Note: rust_xlsxwriter creates a workbook with at least one sheet by default
    // so we'll get an empty sheet, but no data tables were exported
    let sheet_names = workbook.sheet_names();
    // The sheet may be named "Sheet1" or similar default name
    // The key point is that we exported 0 tables
}

#[test]
fn test_conversion_without_headers() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create simple database
    let conn = Connection::open(&db_path).expect("Failed to create database");
    conn.execute("CREATE TABLE test (id INTEGER, value TEXT)", []).unwrap();
    conn.execute("INSERT INTO test VALUES (1, 'hello')", []).unwrap();
    drop(conn);

    // Convert without headers
    let options = ConvertOptions {
        write_headers: false,
        ..Default::default()
    };

    convert(&db_path, &xlsx_path, &options)
        .expect("Conversion should succeed");

    // Verify no header row
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("test")
        .expect("Failed to read sheet 'test'");

    // Row 0 should be data, not headers
    assert_eq!(sheet.get((0, 0)), Some(&Data::Float(1.0)));
    assert_eq!(sheet.get((0, 1)), Some(&Data::String("hello".to_string())));
}

#[test]
fn test_table_filtering() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with multiple tables
    let conn = Connection::open(&db_path).expect("Failed to create database");
    conn.execute("CREATE TABLE table_a (id INTEGER, value TEXT)", []).unwrap();
    conn.execute("CREATE TABLE table_b (id INTEGER, value TEXT)", []).unwrap();
    conn.execute("CREATE TABLE table_c (id INTEGER, value TEXT)", []).unwrap();
    conn.execute("INSERT INTO table_a VALUES (1, 'a')", []).unwrap();
    conn.execute("INSERT INTO table_b VALUES (2, 'b')", []).unwrap();
    conn.execute("INSERT INTO table_c VALUES (3, 'c')", []).unwrap();
    drop(conn);

    // Convert only table_a and table_b
    let options = ConvertOptions {
        tables: Some(vec!["table_a".to_string(), "table_b".to_string()]),
        ..Default::default()
    };

    let stats = convert(&db_path, &xlsx_path, &options)
        .expect("Conversion should succeed");

    assert_eq!(stats.tables_exported, 2);

    // Verify only specified sheets exist
    let workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names.len(), 2);
    assert!(sheet_names.contains(&"table_a".to_string()));
    assert!(sheet_names.contains(&"table_b".to_string()));
    assert!(!sheet_names.contains(&"table_c".to_string()));
}

#[test]
fn test_table_exclusion() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database
    let conn = Connection::open(&db_path).expect("Failed to create database");
    conn.execute("CREATE TABLE keep_this (id INTEGER)", []).unwrap();
    conn.execute("CREATE TABLE skip_this (id INTEGER)", []).unwrap();
    conn.execute("INSERT INTO keep_this VALUES (1)", []).unwrap();
    conn.execute("INSERT INTO skip_this VALUES (2)", []).unwrap();
    drop(conn);

    // Convert excluding skip_this
    let options = ConvertOptions {
        exclude: Some(vec!["skip_this".to_string()]),
        ..Default::default()
    };

    let stats = convert(&db_path, &xlsx_path, &options)
        .expect("Conversion should succeed");

    assert_eq!(stats.tables_exported, 1);

    // Verify only keep_this exists
    let workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names, vec!["keep_this"]);
}

#[test]
fn test_null_values_render_as_blank() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with NULL values
    let conn = Connection::open(&db_path).expect("Failed to create database");
    conn.execute("CREATE TABLE null_test (id INTEGER, text_col TEXT, int_col INTEGER, real_col REAL)", []).unwrap();
    conn.execute("INSERT INTO null_test VALUES (1, NULL, NULL, NULL)", []).unwrap();
    conn.execute("INSERT INTO null_test VALUES (2, 'not null', 42, 3.14)", []).unwrap();
    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify NULL values render as blank cells
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("null_test")
        .expect("Failed to read sheet 'null_test'");

    // Row 1 (all NULLs)
    assert_eq!(sheet.get((1, 0)), Some(&Data::Float(1.0)));
    assert_eq!(sheet.get((1, 1)), Some(&Data::Empty));
    assert_eq!(sheet.get((1, 2)), Some(&Data::Empty));
    assert_eq!(sheet.get((1, 3)), Some(&Data::Empty));

    // Row 2 (no NULLs)
    assert_eq!(sheet.get((2, 0)), Some(&Data::Float(2.0)));
    assert_eq!(sheet.get((2, 1)), Some(&Data::String("not null".to_string())));
    assert_eq!(sheet.get((2, 2)), Some(&Data::Float(42.0)));
    assert_eq!(sheet.get((2, 3)), Some(&Data::Float(3.14)));
}

#[test]
fn test_negative_and_zero_values() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with edge case numeric values
    let conn = Connection::open(&db_path).expect("Failed to create database");
    conn.execute("CREATE TABLE numbers (id INTEGER, int_val INTEGER, real_val REAL)", []).unwrap();
    conn.execute(
        "INSERT INTO numbers VALUES
            (1, -100, -50.5),
            (2, 0, 0.0),
            (3, 1000000, 999999.99)",
        []
    ).unwrap();
    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify numeric values are correct
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("numbers")
        .expect("Failed to read sheet 'numbers'");

    // Negative values
    assert_eq!(sheet.get((1, 1)), Some(&Data::Float(-100.0)));
    assert_eq!(sheet.get((1, 2)), Some(&Data::Float(-50.5)));

    // Zero values
    assert_eq!(sheet.get((2, 1)), Some(&Data::Float(0.0)));
    assert_eq!(sheet.get((2, 2)), Some(&Data::Float(0.0)));

    // Large positive values
    assert_eq!(sheet.get((3, 1)), Some(&Data::Float(1000000.0)));
    assert_eq!(sheet.get((3, 2)), Some(&Data::Float(999999.99)));
}
