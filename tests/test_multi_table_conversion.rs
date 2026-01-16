// Integration tests for multi-table SQLite to XLSX conversion

use calamine::{Reader, Xlsx, open_workbook, Data};
use rusqlite::Connection;
use tempfile::TempDir;
use sqlite_to_xlsx::{convert, ConvertOptions};

#[test]
fn test_multi_table_conversion_all_tables() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with multiple tables
    let conn = Connection::open(&db_path).expect("Failed to create database");

    // Users table
    conn.execute(
        "CREATE TABLE users (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            email TEXT
        )",
        [],
    ).unwrap();
    conn.execute(
        "INSERT INTO users (id, name, email) VALUES
            (1, 'Alice', 'alice@example.com'),
            (2, 'Bob', 'bob@example.com')",
        [],
    ).unwrap();

    // Products table
    conn.execute(
        "CREATE TABLE products (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            price REAL
        )",
        [],
    ).unwrap();
    conn.execute(
        "INSERT INTO products (id, name, price) VALUES
            (1, 'Widget', 19.99),
            (2, 'Gadget', 29.99),
            (3, 'Doohickey', 9.99)",
        [],
    ).unwrap();

    // Orders table
    conn.execute(
        "CREATE TABLE orders (
            id INTEGER PRIMARY KEY,
            user_id INTEGER,
            product_id INTEGER,
            quantity INTEGER
        )",
        [],
    ).unwrap();
    conn.execute(
        "INSERT INTO orders (id, user_id, product_id, quantity) VALUES
            (1, 1, 1, 2),
            (2, 2, 2, 1)",
        [],
    ).unwrap();

    drop(conn);

    // Convert all tables
    let options = ConvertOptions::default();
    let stats = convert(&db_path, &xlsx_path, &options)
        .expect("Conversion should succeed");

    assert_eq!(stats.tables_exported, 3);
    assert_eq!(stats.total_rows, 7); // 2 + 3 + 2

    // Verify XLSX has all sheets
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names.len(), 3);

    // Sheets should be sorted alphabetically
    assert_eq!(sheet_names[0], "orders");
    assert_eq!(sheet_names[1], "products");
    assert_eq!(sheet_names[2], "users");

    // Verify users sheet content
    let users_sheet = workbook.worksheet_range("users")
        .expect("Failed to read sheet 'users'");

    assert_eq!(users_sheet.get((0, 0)), Some(&Data::String("id".to_string())));
    assert_eq!(users_sheet.get((0, 1)), Some(&Data::String("name".to_string())));
    assert_eq!(users_sheet.get((0, 2)), Some(&Data::String("email".to_string())));
    assert_eq!(users_sheet.get((1, 0)), Some(&Data::Float(1.0)));
    assert_eq!(users_sheet.get((1, 1)), Some(&Data::String("Alice".to_string())));
    assert_eq!(users_sheet.get((2, 0)), Some(&Data::Float(2.0)));

    // Verify products sheet
    let products_sheet = workbook.worksheet_range("products")
        .expect("Failed to read sheet 'products'");

    assert_eq!(products_sheet.get((0, 0)), Some(&Data::String("id".to_string())));
    assert_eq!(products_sheet.get((0, 1)), Some(&Data::String("name".to_string())));
    assert_eq!(products_sheet.get((0, 2)), Some(&Data::String("price".to_string())));
    assert_eq!(products_sheet.get((1, 2)), Some(&Data::Float(19.99)));
    assert_eq!(products_sheet.get((2, 2)), Some(&Data::Float(29.99)));

    // Verify orders sheet
    let orders_sheet = workbook.worksheet_range("orders")
        .expect("Failed to read sheet 'orders'");

    assert_eq!(orders_sheet.get((0, 0)), Some(&Data::String("id".to_string())));
    assert_eq!(orders_sheet.get((1, 0)), Some(&Data::Float(1.0)));
}

#[test]
fn test_sheet_name_sanitization_special_characters() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with tables that have invalid Excel sheet name characters
    let conn = Connection::open(&db_path).expect("Failed to create database");

    // Backslashes need to be escaped differently in SQL
    conn.execute("CREATE TABLE \"table/with/slashes\" (id INTEGER)", []).unwrap();
    conn.execute("CREATE TABLE \"table_with_backslashes\" (id INTEGER)", []).unwrap();
    conn.execute("CREATE TABLE \"table:with:colons\" (id INTEGER)", []).unwrap();
    conn.execute("CREATE TABLE \"table*with*asterisks\" (id INTEGER)", []).unwrap();
    conn.execute("CREATE TABLE \"table?with?questions\" (id INTEGER)", []).unwrap();

    conn.execute("INSERT INTO \"table/with/slashes\" VALUES (1)", []).unwrap();
    conn.execute("INSERT INTO \"table_with_backslashes\" VALUES (2)", []).unwrap();
    conn.execute("INSERT INTO \"table:with:colons\" VALUES (3)", []).unwrap();
    conn.execute("INSERT INTO \"table*with*asterisks\" VALUES (4)", []).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify sheet names are sanitized
    let workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();

    // Invalid characters should be replaced with underscores
    assert!(sheet_names.contains(&"table_with_slashes".to_string()));
    assert!(sheet_names.contains(&"table_with_backslashes".to_string()));
    assert!(sheet_names.contains(&"table_with_colons".to_string()));
    assert!(sheet_names.contains(&"table_with_asterisks".to_string()));
    assert!(sheet_names.contains(&"table_with_questions".to_string()));
}

#[test]
fn test_duplicate_sheet_names_get_suffix() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with table names that will collide after sanitization
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute("CREATE TABLE \"table/one\" (id INTEGER)", []).unwrap();
    conn.execute("CREATE TABLE \"table\\one\" (id INTEGER)", []).unwrap();
    conn.execute("CREATE TABLE \"table:one\" (id INTEGER)", []).unwrap();

    conn.execute("INSERT INTO \"table/one\" VALUES (1)", []).unwrap();
    conn.execute("INSERT INTO \"table\\one\" VALUES (2)", []).unwrap();
    conn.execute("INSERT INTO \"table:one\" VALUES (3)", []).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify duplicate names get suffixes
    let workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names.len(), 3);

    // First one stays as-is, others get suffixes
    assert!(sheet_names.contains(&"table_one".to_string()));
    assert!(sheet_names.contains(&"table_one_1".to_string()));
    assert!(sheet_names.contains(&"table_one_2".to_string()));
}

#[test]
fn test_reserved_history_name_handling() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with reserved "History" table name
    let conn = Connection::open(&db_path).expect("Failed to create database");

    // SQLite is case-insensitive for table names, so we need truly different names
    conn.execute("CREATE TABLE History (id INTEGER, event TEXT)", []).unwrap();
    conn.execute("CREATE TABLE old_records (id INTEGER, event TEXT)", []).unwrap();
    conn.execute("CREATE TABLE archive_data (id INTEGER, event TEXT)", []).unwrap();

    conn.execute("INSERT INTO History VALUES (1, 'event1')", []).unwrap();
    conn.execute("INSERT INTO old_records VALUES (2, 'event2')", []).unwrap();
    conn.execute("INSERT INTO archive_data VALUES (3, 'event3')", []).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify "History" is handled (gets underscore to avoid conflict)
    let workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names.len(), 3);

    // History gets suffix because it's reserved
    assert!(sheet_names.contains(&"History_".to_string()));
    // Other names are not reserved and keep their sanitized forms
    assert!(sheet_names.contains(&"old_records".to_string()));
    assert!(sheet_names.contains(&"archive_data".to_string()));
}

#[test]
fn test_very_long_table_names() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with very long table names
    let conn = Connection::open(&db_path).expect("Failed to create database");

    let long_name1 = "this_is_a_very_long_table_name_that_exceeds_the_thirty_one_character_limit";
    let long_name2 = "this_is_another_very_long_table_name_that_exceeds_the_limit";

    conn.execute(&format!("CREATE TABLE \"{}\" (id INTEGER)", long_name1), []).unwrap();
    conn.execute(&format!("CREATE TABLE \"{}\" (id INTEGER)", long_name2), []).unwrap();

    conn.execute(&format!("INSERT INTO \"{}\" VALUES (1)", long_name1), []).unwrap();
    conn.execute(&format!("INSERT INTO \"{}\" VALUES (2)", long_name2), []).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify names are truncated to 31 chars
    let workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names.len(), 2);

    // Both should be truncated to 28 chars to leave room for suffix
    assert!(sheet_names[0].len() <= 31);
    assert!(sheet_names[1].len() <= 31);

    // First one gets base name (truncated to 28 chars)
    assert_eq!(sheet_names[0], "this_is_a_very_long_table_na");
    assert_eq!(sheet_names[0].len(), 28);

    // Second one gets different base name (truncated to 28 chars)
    // Since the truncated names are different, no suffix is added
    assert_eq!(sheet_names[1], "this_is_another_very_long_ta");
    assert_eq!(sheet_names[1].len(), 28);
}

#[test]
fn test_tables_with_different_schemas() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create tables with different schemas
    let conn = Connection::open(&db_path).expect("Failed to create database");

    // Simple table
    conn.execute("CREATE TABLE simple (id INTEGER)", []).unwrap();
    conn.execute("INSERT INTO simple VALUES (1)", []).unwrap();

    // Wide table (many columns)
    conn.execute(
        "CREATE TABLE wide (
            col1 INTEGER, col2 INTEGER, col3 INTEGER, col4 INTEGER, col5 INTEGER,
            col6 INTEGER, col7 INTEGER, col8 INTEGER, col9 INTEGER, col10 INTEGER
        )",
        [],
    ).unwrap();
    conn.execute(
        "INSERT INTO wide VALUES (1,2,3,4,5,6,7,8,9,10)",
        [],
    ).unwrap();

    // Table with various types
    conn.execute(
        "CREATE TABLE mixed_types (
            int_col INTEGER,
            real_col REAL,
            text_col TEXT,
            blob_col BLOB
        )",
        [],
    ).unwrap();
    conn.execute(
        "INSERT INTO mixed_types VALUES (42, 3.14, 'hello', x'48656c6c6f')",
        [],
    ).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify all tables are exported with correct structure
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names.len(), 3);

    // Verify wide table has all columns
    let wide_sheet = workbook.worksheet_range("wide")
        .expect("Failed to read sheet 'wide'");

    // Should have 10 columns
    assert_eq!(wide_sheet.get((0, 0)), Some(&Data::String("col1".to_string())));
    assert_eq!(wide_sheet.get((0, 9)), Some(&Data::String("col10".to_string())));
    assert_eq!(wide_sheet.get((1, 0)), Some(&Data::Float(1.0)));
    assert_eq!(wide_sheet.get((1, 9)), Some(&Data::Float(10.0)));
}

#[test]
fn test_blob_handling_strategies_across_tables() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create tables with BLOB data
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute("CREATE TABLE blob_placeholder (id INTEGER, data BLOB)", []).unwrap();
    conn.execute("CREATE TABLE blob_hex (id INTEGER, data BLOB)", []).unwrap();
    conn.execute("CREATE TABLE blob_base64 (id INTEGER, data BLOB)", []).unwrap();
    conn.execute("CREATE TABLE blob_skip (id INTEGER, data BLOB)", []).unwrap();

    let blob_data = vec![0x48, 0x65, 0x6c, 0x6c, 0x6f]; // "Hello"

    for table in &["blob_placeholder", "blob_hex", "blob_base64", "blob_skip"] {
        conn.execute(
            &format!("INSERT INTO {} VALUES (1, ?)", table),
            [blob_data.as_slice()],
        ).unwrap();
    }

    drop(conn);

    // Test with base64 encoding
    let options = ConvertOptions {
        blob_handling: sqlite_to_xlsx::BlobHandling::Base64,
        ..Default::default()
    };

    convert(&db_path, &xlsx_path, &options)
        .expect("Conversion should succeed");

    // Verify BLOB is encoded as base64
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("blob_base64")
        .expect("Failed to read sheet 'blob_base64'");

    // Should be base64 encoded "SGVsbG8="
    let cell_value = sheet.get((1, 1));
    match cell_value {
        Some(Data::String(s)) => assert_eq!(s, "SGVsbG8="),
        _ => panic!("Expected base64 encoded string, got {:?}", cell_value),
    }
}

#[test]
fn test_empty_tables_are_included() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with empty and non-empty tables
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute("CREATE TABLE empty_table (id INTEGER, name TEXT)", []).unwrap();
    conn.execute("CREATE TABLE non_empty (id INTEGER, value TEXT)", []).unwrap();
    conn.execute("INSERT INTO non_empty VALUES (1, 'data')", []).unwrap();

    drop(conn);

    let stats = convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Both tables should be exported
    assert_eq!(stats.tables_exported, 2);
    assert_eq!(stats.total_rows, 1); // Only one row from non_empty table

    // Verify empty table has headers but no data rows
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let empty_sheet = workbook.worksheet_range("empty_table")
        .expect("Failed to read sheet 'empty_table'");

    // Should have headers
    assert_eq!(empty_sheet.get((0, 0)), Some(&Data::String("id".to_string())));
    assert_eq!(empty_sheet.get((0, 1)), Some(&Data::String("name".to_string())));

    // Should have no data rows (only headers = row 0)
    // Check that row 1 doesn't exist (returns None)
    assert_eq!(empty_sheet.get((1, 0)), None);
}

#[test]
fn test_tables_sorted_alphabetically() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create tables in non-alphabetical order
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute("CREATE TABLE zebra (id INTEGER)", []).unwrap();
    conn.execute("CREATE TABLE apple (id INTEGER)", []).unwrap();
    conn.execute("CREATE TABLE banana (id INTEGER)", []).unwrap();

    conn.execute("INSERT INTO zebra VALUES (1)", []).unwrap();
    conn.execute("INSERT INTO apple VALUES (2)", []).unwrap();
    conn.execute("INSERT INTO banana VALUES (3)", []).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify sheets are in alphabetical order
    let workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names, vec!["apple", "banana", "zebra"]);
}
