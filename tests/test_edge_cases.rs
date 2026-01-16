// Integration tests for edge cases in SQLite to XLSX conversion

use calamine::{Reader, Xlsx, open_workbook, Data};
use rusqlite::Connection;
use tempfile::TempDir;
use sqlite_to_xlsx::{convert, ConvertOptions, BlobHandling};

#[test]
fn test_mixed_types_in_same_column() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // SQLite allows mixed types in the same column
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE mixed_types (
            id INTEGER,
            value ANY
        )",
        [],
    ).unwrap();

    // Insert different types in the same column
    conn.execute("INSERT INTO mixed_types VALUES (1, 42)", []).unwrap();           // INTEGER
    conn.execute("INSERT INTO mixed_types VALUES (2, 1.5)", []).unwrap();         // REAL
    conn.execute("INSERT INTO mixed_types VALUES (3, 'hello')", []).unwrap();      // TEXT
    conn.execute("INSERT INTO mixed_types VALUES (4, NULL)", []).unwrap();         // NULL
    conn.execute("INSERT INTO mixed_types VALUES (5, x'48656c6c6f')", []).unwrap(); // BLOB

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify all types are correctly converted
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("mixed_types")
        .expect("Failed to read sheet 'mixed_types'");

    // Row 1: INTEGER
    assert_eq!(sheet.get((1, 0)), Some(&Data::Float(1.0)));
    assert_eq!(sheet.get((1, 1)), Some(&Data::Float(42.0)));

    // Row 2: REAL
    assert_eq!(sheet.get((2, 1)), Some(&Data::Float(1.5)));

    // Row 3: TEXT
    assert_eq!(sheet.get((3, 1)), Some(&Data::String("hello".to_string())));

    // Row 4: NULL
    assert_eq!(sheet.get((4, 1)), Some(&Data::Empty));

    // Row 5: BLOB (placeholder by default)
    match sheet.get((5, 1)) {
        Some(Data::String(s)) => assert!(s.starts_with("[BLOB:")),
        _ => panic!("Expected BLOB placeholder string"),
    }
}

#[test]
fn test_special_float_values() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test special float values
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE special_floats (
            id INTEGER,
            value REAL
        )",
        [],
    ).unwrap();

    // Insert special float values
    conn.execute("INSERT INTO special_floats VALUES (1, 0.0)", []).unwrap();
    conn.execute("INSERT INTO special_floats VALUES (2, -0.0)", []).unwrap();
    // Note: SQLite doesn't natively store NaN/Infinity, but we can test the conversion logic

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("special_floats")
        .expect("Failed to read sheet 'special_floats'");

    assert_eq!(sheet.get((1, 1)), Some(&Data::Float(0.0)));
    assert_eq!(sheet.get((2, 1)), Some(&Data::Float(-0.0)));
}

#[test]
fn test_string_with_quotes_and_newlines() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test strings with special characters
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE special_strings (
            id INTEGER,
            text TEXT
        )",
        [],
    ).unwrap();

    let special_strings = [
        "Line 1\nLine 2\nLine 3",           // Newlines
        "Text with \"quotes\" inside",       // Double quotes
        "Text with 'apostrophes' inside",    // Single quotes
        "Text with both \"double' and 'single\"", // Mixed
        "Tab\there",                         // Tab character
        "Text with \r\n carriage return",    // CRLF
        "Path: C:\\Users\\test",            // Backslashes
        "Formula-like: =SUM(A1:A10)",       // Formula-like text
    ];

    for (i, text) in special_strings.iter().enumerate() {
        conn.execute(
            "INSERT INTO special_strings VALUES (?1, ?2)",
            [&((i as i64) + 1) as &dyn rusqlite::ToSql, &&**text as &dyn rusqlite::ToSql],
        ).unwrap();
    }

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("special_strings")
        .expect("Failed to read sheet 'special_strings'");

    for (i, expected_text) in special_strings.iter().enumerate() {
        let row = ((i + 1) as u32) as usize;
        match sheet.get((row, 1)) {
            Some(Data::String(s)) => {
                // Excel converts \r to _x000D_ in XLSX format (XML escape for carriage return)
                // When we write \r\n, Excel converts it to _x000D_\n
                let expected = if expected_text.contains("\r\n") {
                    expected_text.replace("\r\n", "_x000D_\n")
                } else if expected_text.contains("\r") {
                    expected_text.replace("\r", "_x000D_")
                } else {
                    expected_text.to_string()
                };
                assert_eq!(&**s, expected, "Text mismatch at row {}", row);
            },
            _ => panic!("Expected string at row {}, got {:?}", row, sheet.get((row, 1))),
        }
    }
}

#[test]
fn test_very_long_column_names() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test very long column names
    let conn = Connection::open(&db_path).expect("Failed to create database");

    let long_col_name = "this_is_a_very_long_column_name_that_exceeds_normal_limits_and_keeps_going_on_and_on";

    conn.execute(
        &format!("CREATE TABLE long_cols (id INTEGER, \"{}\" INTEGER)", long_col_name),
        [],
    ).unwrap();

    conn.execute(
        "INSERT INTO long_cols VALUES (1, 42)",
        [],
    ).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("long_cols")
        .expect("Failed to read sheet 'long_cols'");

    // Long column names should be preserved in XLSX
    assert_eq!(sheet.get((0, 0)), Some(&Data::String("id".to_string())));
    assert_eq!(sheet.get((0, 1)), Some(&Data::String(long_col_name.to_string())));
}

#[test]
fn test_table_with_no_primary_key() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test table without primary key
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE no_pk (
            name TEXT,
            age INTEGER,
            city TEXT
        )",
        [],
    ).unwrap();

    conn.execute("INSERT INTO no_pk VALUES ('Alice', 30, 'NYC')", []).unwrap();
    conn.execute("INSERT INTO no_pk VALUES ('Bob', 25, 'LA')", []).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("no_pk")
        .expect("Failed to read sheet 'no_pk'");

    // Headers should be correct without PK
    assert_eq!(sheet.get((0, 0)), Some(&Data::String("name".to_string())));
    assert_eq!(sheet.get((0, 1)), Some(&Data::String("age".to_string())));
    assert_eq!(sheet.get((0, 2)), Some(&Data::String("city".to_string())));

    // Data should be correct
    assert_eq!(sheet.get((1, 0)), Some(&Data::String("Alice".to_string())));
    assert_eq!(sheet.get((2, 0)), Some(&Data::String("Bob".to_string())));
}

#[test]
fn test_composite_primary_key() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test table with composite primary key
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE composite_pk (
            user_id INTEGER,
            item_id INTEGER,
            quantity INTEGER,
            PRIMARY KEY (user_id, item_id)
        )",
        [],
    ).unwrap();

    conn.execute("INSERT INTO composite_pk VALUES (1, 100, 5)", []).unwrap();
    conn.execute("INSERT INTO composite_pk VALUES (1, 101, 3)", []).unwrap();
    conn.execute("INSERT INTO composite_pk VALUES (2, 100, 10)", []).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("composite_pk")
        .expect("Failed to read sheet 'composite_pk'");

    // Verify data is correct
    assert_eq!(sheet.get((1, 0)), Some(&Data::Float(1.0)));
    assert_eq!(sheet.get((1, 1)), Some(&Data::Float(100.0)));
    assert_eq!(sheet.get((1, 2)), Some(&Data::Float(5.0)));
}

#[test]
fn test_all_blob_handling_modes() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");

    let blob_data = vec![0x48, 0x65, 0x6c, 0x6c, 0x6f]; // "Hello"

    // Test Placeholder mode
    {
        let db_path = temp_dir.path().join("placeholder.db");
        let xlsx_path = temp_dir.path().join("placeholder.xlsx");
        let conn = Connection::open(&db_path).unwrap();
        conn.execute("CREATE TABLE test (id INTEGER, data BLOB)", []).unwrap();
        conn.execute("INSERT INTO test VALUES (1, ?)", [blob_data.as_slice()]).unwrap();
        drop(conn);

        let options = ConvertOptions {
            blob_handling: BlobHandling::Placeholder,
            quiet: true,
            ..Default::default()
        };

        convert(&db_path, &xlsx_path, &options).unwrap();

        let mut workbook: Xlsx<_> = open_workbook(&xlsx_path).unwrap();
        let sheet = workbook.worksheet_range("test").expect("Failed to read sheet 'test'");

        match sheet.get((1, 1)) {
            Some(Data::String(s)) => assert_eq!(s, "[BLOB: 5 bytes]"),
            _ => panic!("Expected placeholder string"),
        }
    }

    // Test Hex mode
    {
        let db_path = temp_dir.path().join("hex.db");
        let xlsx_path = temp_dir.path().join("hex.xlsx");
        let conn = Connection::open(&db_path).unwrap();
        conn.execute("CREATE TABLE test (id INTEGER, data BLOB)", []).unwrap();
        conn.execute("INSERT INTO test VALUES (1, ?)", [blob_data.as_slice()]).unwrap();
        drop(conn);

        let options = ConvertOptions {
            blob_handling: BlobHandling::Hex,
            quiet: true,
            ..Default::default()
        };

        convert(&db_path, &xlsx_path, &options).unwrap();

        let mut workbook: Xlsx<_> = open_workbook(&xlsx_path).unwrap();
        let sheet = workbook.worksheet_range("test").expect("Failed to read sheet 'test'");

        match sheet.get((1, 1)) {
            Some(Data::String(s)) => assert_eq!(s, "0x48656c6c6f"),
            _ => panic!("Expected hex string"),
        }
    }

    // Test Base64 mode
    {
        let db_path = temp_dir.path().join("base64.db");
        let xlsx_path = temp_dir.path().join("base64.xlsx");
        let conn = Connection::open(&db_path).unwrap();
        conn.execute("CREATE TABLE test (id INTEGER, data BLOB)", []).unwrap();
        conn.execute("INSERT INTO test VALUES (1, ?)", [blob_data.as_slice()]).unwrap();
        drop(conn);

        let options = ConvertOptions {
            blob_handling: BlobHandling::Base64,
            quiet: true,
            ..Default::default()
        };

        convert(&db_path, &xlsx_path, &options).unwrap();

        let mut workbook: Xlsx<_> = open_workbook(&xlsx_path).unwrap();
        let sheet = workbook.worksheet_range("test").expect("Failed to read sheet 'test'");

        match sheet.get((1, 1)) {
            Some(Data::String(s)) => assert_eq!(s, "SGVsbG8="),
            _ => panic!("Expected base64 string"),
        }
    }

    // Test Skip mode
    {
        let db_path = temp_dir.path().join("skip.db");
        let xlsx_path = temp_dir.path().join("skip.xlsx");
        let conn = Connection::open(&db_path).unwrap();
        conn.execute("CREATE TABLE test (id INTEGER, data BLOB)", []).unwrap();
        conn.execute("INSERT INTO test VALUES (1, ?)", [blob_data.as_slice()]).unwrap();
        drop(conn);

        let options = ConvertOptions {
            blob_handling: BlobHandling::Skip,
            quiet: true,
            ..Default::default()
        };

        convert(&db_path, &xlsx_path, &options).unwrap();

        let mut workbook: Xlsx<_> = open_workbook(&xlsx_path).unwrap();
        let sheet = workbook.worksheet_range("test").expect("Failed to read sheet 'test'");

        // Should be empty (blank cell)
        assert_eq!(sheet.get((1, 1)), Some(&Data::Empty));
    }
}

#[test]
fn test_empty_strings_vs_null() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test distinction between empty string and NULL
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE empty_vs_null (
            id INTEGER,
            empty_str TEXT,
            null_val TEXT,
            normal_str TEXT
        )",
        [],
    ).unwrap();

    conn.execute(
        "INSERT INTO empty_vs_null VALUES (1, '', NULL, 'normal')",
        [],
    ).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("empty_vs_null")
        .expect("Failed to read sheet 'empty_vs_null'");

    // Empty string is stored as Empty in calamine (Excel doesn't distinguish empty from null)
    assert_eq!(sheet.get((1, 1)), Some(&Data::Empty));

    // NULL should be Empty
    assert_eq!(sheet.get((1, 2)), Some(&Data::Empty));

    // Normal string
    assert_eq!(sheet.get((1, 3)), Some(&Data::String("normal".to_string())));
}

#[test]
fn test_single_column_table() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test table with only one column
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute("CREATE TABLE single_col (value TEXT)", []).unwrap();
    conn.execute("INSERT INTO single_col VALUES ('a')", []).unwrap();
    conn.execute("INSERT INTO single_col VALUES ('b')", []).unwrap();
    conn.execute("INSERT INTO single_col VALUES ('c')", []).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("single_col")
        .expect("Failed to read sheet 'single_col'");

    assert_eq!(sheet.get((0, 0)), Some(&Data::String("value".to_string())));
    assert_eq!(sheet.get((1, 0)), Some(&Data::String("a".to_string())));
    assert_eq!(sheet.get((2, 0)), Some(&Data::String("b".to_string())));
    assert_eq!(sheet.get((3, 0)), Some(&Data::String("c".to_string())));
}

#[test]
fn test_integer_boundary_values() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test integer boundary values
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE int_boundaries (
            id INTEGER,
            value INTEGER
        )",
        [],
    ).unwrap();

    // Test various integer values
    let test_values = [
        0_i64,
        1,
        -1,
        127,
        -128,
        255,
        -256,
        32767,
        -32768,
        65536,
        -65536,
        2147483647,      // i32::MAX
        -2147483648_i64,  // i32::MIN
        9007199254740991, // Max safe integer (2^53 - 1)
        -9007199254740991, // -(2^53 - 1)
    ];

    for (i, val) in test_values.iter().enumerate() {
        conn.execute(
            "INSERT INTO int_boundaries VALUES (?1, ?2)",
            [&((i as i64) + 1) as &dyn rusqlite::ToSql, val as &dyn rusqlite::ToSql],
        ).unwrap();
    }

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions {
        quiet: true,
        ..Default::default()
    }).expect("Conversion should succeed");

    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("int_boundaries")
        .expect("Failed to read sheet 'int_boundaries'");

    // Verify values are correctly preserved
    for (i, expected_val) in test_values.iter().enumerate() {
        let row = ((i + 1) as u32) as usize;
        match sheet.get((row, 1)) {
            Some(Data::Float(v)) => {
                assert_eq!(*v, *expected_val as f64, "Float mismatch at row {}", row);
            },
            Some(Data::String(s)) => {
                // Large integers may be stored as strings
                assert_eq!(s.parse::<i64>().unwrap(), *expected_val, "String int mismatch at row {}", row);
            },
            _ => panic!("Unexpected value type at row {}: {:?}", row, sheet.get((row, 1))),
        }
    }
}

#[test]
fn test_sqlite_reserved_words_as_column_names() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test SQLite reserved words as column names
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE reserved_words (
            \"select\" TEXT,
            \"insert\" TEXT,
            \"update\" TEXT,
            \"delete\" TEXT,
            \"drop\" TEXT
        )",
        [],
    ).unwrap();

    conn.execute(
        "INSERT INTO reserved_words VALUES ('a', 'b', 'c', 'd', 'e')",
        [],
    ).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("reserved_words")
        .expect("Failed to read sheet 'reserved_words'");

    // Verify reserved word column names are preserved
    assert_eq!(sheet.get((0, 0)), Some(&Data::String("select".to_string())));
    assert_eq!(sheet.get((0, 1)), Some(&Data::String("insert".to_string())));
    assert_eq!(sheet.get((0, 2)), Some(&Data::String("update".to_string())));
    assert_eq!(sheet.get((0, 3)), Some(&Data::String("delete".to_string())));
    assert_eq!(sheet.get((0, 4)), Some(&Data::String("drop".to_string())));
}
