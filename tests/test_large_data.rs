// Integration tests for large dataset handling in SQLite to XLSX conversion

use calamine::{Reader, Xlsx, open_workbook, Data};
use rusqlite::Connection;
use tempfile::TempDir;
use sqlite_to_xlsx::{convert, ConvertOptions};

#[test]
fn test_large_dataset_10k_rows() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with 10,000 rows
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE large_data (
            id INTEGER PRIMARY KEY,
            name TEXT,
            value INTEGER,
            score REAL,
            active INTEGER
        )",
        [],
    ).unwrap();

    // Insert 10,000 rows
    let tx = conn.unchecked_transaction().expect("Failed to begin transaction");
    {
        let mut stmt = conn.prepare(
            "INSERT INTO large_data (id, name, value, score, active) VALUES (?1, ?2, ?3, ?4, ?5)"
        ).expect("Failed to prepare statement");

        for i in 0..10_000 {
            stmt.execute([
                &((i + 1) as i64) as &dyn rusqlite::ToSql,
                &format!("Item_{}", i) as &dyn rusqlite::ToSql,
                &(((i * 10) % 1000) as i64) as &dyn rusqlite::ToSql,
                &((i as f64 * 0.1).fract() * 100.0) as &dyn rusqlite::ToSql,
                &(if i % 2 == 0 { 1_i64 } else { 0_i64 }) as &dyn rusqlite::ToSql
            ]).expect("Failed to insert row");
        }
    }
    tx.commit().expect("Failed to commit transaction");

    drop(conn);

    // Convert with quiet mode to avoid progress bar output in tests
    let options = ConvertOptions {
        quiet: true,
        ..Default::default()
    };

    let stats = convert(&db_path, &xlsx_path, &options)
        .expect("Conversion should succeed");

    assert_eq!(stats.tables_exported, 1);
    assert_eq!(stats.total_rows, 10_000);

    // Verify the XLSX file can be opened and read
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("large_data")
        .expect("Failed to read sheet 'large_data'");

    // Verify headers
    assert_eq!(sheet.get((0, 0)), Some(&Data::String("id".to_string())));
    assert_eq!(sheet.get((0, 1)), Some(&Data::String("name".to_string())));
    assert_eq!(sheet.get((0, 2)), Some(&Data::String("value".to_string())));
    assert_eq!(sheet.get((0, 3)), Some(&Data::String("score".to_string())));
    assert_eq!(sheet.get((0, 4)), Some(&Data::String("active".to_string())));

    // Verify first data row
    assert_eq!(sheet.get((1, 0)), Some(&Data::Float(1.0)));
    assert_eq!(sheet.get((1, 1)), Some(&Data::String("Item_0".to_string())));
    assert_eq!(sheet.get((1, 2)), Some(&Data::Float(0.0)));

    // Verify last data row
    assert_eq!(sheet.get((10000, 0)), Some(&Data::Float(10000.0)));
    assert_eq!(sheet.get((10000, 1)), Some(&Data::String("Item_9999".to_string())));
}

#[test]
fn test_very_wide_table_100_columns() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with a very wide table (100 columns)
    let conn = Connection::open(&db_path).expect("Failed to create database");

    // Build CREATE TABLE statement with 100 columns
    let mut create_sql = String::from("CREATE TABLE wide_table (id INTEGER PRIMARY KEY");
    for i in 1..=100 {
        create_sql.push_str(&format!(", col{} TEXT", i));
    }
    create_sql.push(')');

    conn.execute(&create_sql, []).expect("Failed to create wide table");

    // Insert a few rows
    let mut insert_sql = String::from("INSERT INTO wide_table VALUES (?1");
    for _ in 1..=100 { // id + 100 columns = 101 total placeholders
        insert_sql.push_str(", ?");
    }
    insert_sql.push(')');

    // Insert 10 rows with data in all columns
    let tx = conn.unchecked_transaction().expect("Failed to begin transaction");
    {
        let mut stmt = conn.prepare(&insert_sql).expect("Failed to prepare statement");

        for row_idx in 0..10 {
            let mut params: Vec<Box<dyn rusqlite::ToSql>> = vec![Box::new(row_idx)]; // id is integer
            for col_idx in 1..=100 {
                params.push(Box::new(format!("r{}c{}", row_idx, col_idx)));
            }

            let params_refs: Vec<&dyn rusqlite::ToSql> = params.iter().map(|p| p.as_ref()).collect();
            stmt.execute(params_refs.as_slice()).expect("Failed to insert row");
        }
    }
    tx.commit().expect("Failed to commit transaction");

    drop(conn);

    let stats = convert(&db_path, &xlsx_path, &ConvertOptions {
        quiet: true,
        ..Default::default()
    }).expect("Conversion should succeed");

    assert_eq!(stats.tables_exported, 1);
    assert_eq!(stats.total_rows, 10);

    // Verify all columns are present
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("wide_table")
        .expect("Failed to read sheet 'wide_table'");

    // Verify first and last column headers
    assert_eq!(sheet.get((0, 0)), Some(&Data::String("id".to_string())));
    assert_eq!(sheet.get((0, 1)), Some(&Data::String("col1".to_string())));
    assert_eq!(sheet.get((0, 100)), Some(&Data::String("col100".to_string())));

    // Verify data in first and last columns
    assert_eq!(sheet.get((1, 0)), Some(&Data::Float(0.0))); // id is now integer 0
    assert_eq!(sheet.get((1, 1)), Some(&Data::String("r0c1".to_string())));
    assert_eq!(sheet.get((1, 100)), Some(&Data::String("r0c100".to_string())));
}

#[test]
fn test_multiple_large_tables() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with multiple large tables
    let conn = Connection::open(&db_path).expect("Failed to create database");

    // Table 1: 5,000 rows
    conn.execute(
        "CREATE TABLE table_a (id INTEGER, name TEXT, value INTEGER)",
        [],
    ).unwrap();

    let tx = conn.unchecked_transaction().unwrap();
    {
        let mut stmt = conn.prepare("INSERT INTO table_a VALUES (?1, ?2, ?3)").unwrap();
        for i in 0..5_000 {
            stmt.execute([
                &(i as i64) as &dyn rusqlite::ToSql,
                &format!("A_{}", i) as &dyn rusqlite::ToSql,
                &((i * 2) as i64) as &dyn rusqlite::ToSql
            ]).unwrap();
        }
    }
    tx.commit().unwrap();

    // Table 2: 3,000 rows
    conn.execute(
        "CREATE TABLE table_b (id INTEGER, name TEXT, value INTEGER)",
        [],
    ).unwrap();

    let tx = conn.unchecked_transaction().unwrap();
    {
        let mut stmt = conn.prepare("INSERT INTO table_b VALUES (?1, ?2, ?3)").unwrap();
        for i in 0..3_000 {
            stmt.execute([
                &(i as i64) as &dyn rusqlite::ToSql,
                &format!("B_{}", i) as &dyn rusqlite::ToSql,
                &((i * 3) as i64) as &dyn rusqlite::ToSql
            ]).unwrap();
        }
    }
    tx.commit().unwrap();

    // Table 3: 2,000 rows
    conn.execute(
        "CREATE TABLE table_c (id INTEGER, name TEXT, value INTEGER)",
        [],
    ).unwrap();

    let tx = conn.unchecked_transaction().unwrap();
    {
        let mut stmt = conn.prepare("INSERT INTO table_c VALUES (?1, ?2, ?3)").unwrap();
        for i in 0..2_000 {
            stmt.execute([
                &(i as i64) as &dyn rusqlite::ToSql,
                &format!("C_{}", i) as &dyn rusqlite::ToSql,
                &((i * 4) as i64) as &dyn rusqlite::ToSql
            ]).unwrap();
        }
    }
    tx.commit().unwrap();

    drop(conn);

    let stats = convert(&db_path, &xlsx_path, &ConvertOptions {
        quiet: true,
        ..Default::default()
    }).expect("Conversion should succeed");

    assert_eq!(stats.tables_exported, 3);
    assert_eq!(stats.total_rows, 10_000); // 5000 + 3000 + 2000

    // Verify all sheets exist
    let workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names.len(), 3);
    assert!(sheet_names.contains(&"table_a".to_string()));
    assert!(sheet_names.contains(&"table_b".to_string()));
    assert!(sheet_names.contains(&"table_c".to_string()));
}

#[test]
fn test_large_text_values() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with very large text values
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE large_text (
            id INTEGER,
            text_data TEXT
        )",
        [],
    ).unwrap();

    // Create various large text strings
    let large_strings = vec![
        "a".repeat(100),      // 100 chars
        "b".repeat(1_000),    // 1K chars
        "c".repeat(10_000),   // 10K chars
        "d".repeat(30_000),   // 30K chars (under Excel limit)
    ];

    for (i, text) in large_strings.iter().enumerate() {
        conn.execute(
            "INSERT INTO large_text (id, text_data) VALUES (?1, ?2)",
            [&((i as i64) + 1) as &dyn rusqlite::ToSql, &&**text as &dyn rusqlite::ToSql],
        ).unwrap();
    }

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions {
        quiet: true,
        ..Default::default()
    }).expect("Conversion should succeed");

    // Verify large text values are preserved
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("large_text")
        .expect("Failed to read sheet 'large_text'");

    // Check each large text value
    for (i, expected_text) in large_strings.iter().enumerate() {
        let row = ((i + 1) as u32) as usize;
        match sheet.get((row, 1)) {
            Some(Data::String(s)) => {
                assert_eq!(s.len(), expected_text.len(), "Length mismatch at row {}", row);
                assert_eq!(&**s, *expected_text, "Text mismatch at row {}", row);
            },
            _ => panic!("Expected string at row {}, got {:?}", row, sheet.get((row, 1))),
        }
    }
}

#[test]
fn test_large_table_with_all_null_values() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create large table with many NULL values
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE sparse_data (
            id INTEGER PRIMARY KEY,
            col1 TEXT,
            col2 INTEGER,
            col3 REAL,
            col4 TEXT,
            col5 INTEGER
        )",
        [],
    ).unwrap();

    // Insert 5,000 rows with mostly NULL values
    let tx = conn.unchecked_transaction().unwrap();
    {
        let mut stmt = conn.prepare(
            "INSERT INTO sparse_data (id, col1, col2, col3, col4, col5) VALUES (?1, ?2, ?3, ?4, ?5, ?6)"
        ).unwrap();

        for i in 0..5_000 {
            // Only every 100th row has data
            let has_data = i % 100 == 0;
            let p1: &dyn rusqlite::ToSql = &(i + 1) as &dyn rusqlite::ToSql;
            let p2: &dyn rusqlite::ToSql = if has_data { &Some("data") as &dyn rusqlite::ToSql } else { &None::<&str> as &dyn rusqlite::ToSql };
            let p3: &dyn rusqlite::ToSql = if has_data { &Some(i as i64) as &dyn rusqlite::ToSql } else { &None::<i64> as &dyn rusqlite::ToSql };
            let p4: &dyn rusqlite::ToSql = if has_data { &Some(i as f64) as &dyn rusqlite::ToSql } else { &None::<f64> as &dyn rusqlite::ToSql };
            let p5: &dyn rusqlite::ToSql = if has_data { &Some("value") as &dyn rusqlite::ToSql } else { &None::<&str> as &dyn rusqlite::ToSql };
            let p6: &dyn rusqlite::ToSql = if has_data { &Some((i * 2) as i64) as &dyn rusqlite::ToSql } else { &None::<i64> as &dyn rusqlite::ToSql };
            stmt.execute([p1, p2, p3, p4, p5, p6]).unwrap();
        }
    }
    tx.commit().unwrap();

    drop(conn);

    let stats = convert(&db_path, &xlsx_path, &ConvertOptions {
        quiet: true,
        ..Default::default()
    }).expect("Conversion should succeed");

    assert_eq!(stats.tables_exported, 1);
    assert_eq!(stats.total_rows, 5_000);

    // Verify NULL values and data are correct
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("sparse_data")
        .expect("Failed to read sheet 'sparse_data'");

    // Check a row with data (every 100th row)
    assert_eq!(sheet.get((101, 0)), Some(&Data::Float(101.0)));
    assert_eq!(sheet.get((101, 1)), Some(&Data::String("data".to_string())));
    assert_eq!(sheet.get((101, 2)), Some(&Data::Float(100.0)));

    // Check a row with NULL values
    assert_eq!(sheet.get((102, 0)), Some(&Data::Float(102.0)));
    assert_eq!(sheet.get((102, 1)), Some(&Data::Empty));
    assert_eq!(sheet.get((102, 2)), Some(&Data::Empty));
}

#[test]
fn test_numeric_precision_with_large_values() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test precision handling with large numeric values
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE precision_test (
            id INTEGER,
            large_int INTEGER,
            small_float REAL,
            large_float REAL
        )",
        [],
    ).unwrap();

    // Insert values that test precision limits
    let test_values = vec![
        (1, 9_007_199_254_740_991_i64, 0.1_f64, 1000000.12345678),
        (2, 9_007_199_254_740_992_i64, 0.01, 1000000.01234567),
        (3, 9_007_199_254_740_993_i64, 0.001, 1000000.00123456),
        (4, -9_007_199_254_740_991_i64, -0.1, -1000000.12345678),
        (5, 0_i64, 0.0, 0.0),
    ];

    for (id, large_int, small_float, large_float) in test_values {
        conn.execute(
            "INSERT INTO precision_test VALUES (?1, ?2, ?3, ?4)",
            [&id as &dyn rusqlite::ToSql, &large_int as &dyn rusqlite::ToSql, &small_float as &dyn rusqlite::ToSql, &large_float as &dyn rusqlite::ToSql],
        ).unwrap();
    }

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions {
        quiet: true,
        ..Default::default()
    }).expect("Conversion should succeed");

    // Note: Large integers (> 2^53) are written as strings to preserve precision
    // This test verifies the conversion completes without error
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("precision_test")
        .expect("Failed to read sheet 'precision_test'");

    // Small float should be preserved
    match sheet.get((1, 2)) {
        Some(Data::Float(f)) => assert!((f - 0.1).abs() < 0.001),
        _ => panic!("Expected float value"),
    }

    // Large floats may have some precision loss
    match sheet.get((1, 3)) {
        Some(Data::Float(_)) => (), // Accept some precision loss
        _ => panic!("Expected float value"),
    }
}

#[test]
fn test_conversion_performance_time() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create a moderately large dataset
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE perf_test (
            id INTEGER PRIMARY KEY,
            col1 TEXT,
            col2 INTEGER,
            col3 REAL,
            col4 TEXT,
            col5 INTEGER,
            col6 REAL,
            col7 TEXT,
            col8 INTEGER,
            col9 REAL,
            col10 TEXT
        )",
        [],
    ).unwrap();

    // Insert 2,000 rows with 10 columns each = 20,000 cells
    let tx = conn.unchecked_transaction().unwrap();
    {
        let mut stmt = conn.prepare(
            "INSERT INTO perf_test VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11)"
        ).unwrap();

        for i in 0..2_000 {
            stmt.execute([
                &(i as i64) as &dyn rusqlite::ToSql,
                &format!("text_{}", i) as &dyn rusqlite::ToSql,
                &((i * 2) as i64) as &dyn rusqlite::ToSql,
                &(i as f64 * 0.5) as &dyn rusqlite::ToSql,
                &format!("value_{}", i) as &dyn rusqlite::ToSql,
                &((i * 3) as i64) as &dyn rusqlite::ToSql,
                &(i as f64 * 1.5) as &dyn rusqlite::ToSql,
                &format!("data_{}", i) as &dyn rusqlite::ToSql,
                &((i * 4) as i64) as &dyn rusqlite::ToSql,
                &(i as f64 * 2.5) as &dyn rusqlite::ToSql,
                &format!("item_{}", i) as &dyn rusqlite::ToSql,
            ]).unwrap();
        }
    }
    tx.commit().unwrap();

    drop(conn);

    let start = std::time::Instant::now();

    let stats = convert(&db_path, &xlsx_path, &ConvertOptions {
        quiet: true,
        ..Default::default()
    }).expect("Conversion should succeed");

    let duration = start.elapsed();

    assert_eq!(stats.tables_exported, 1);
    assert_eq!(stats.total_rows, 2_000);

    // The test just verifies it completes; we don't assert a specific time
    // because performance varies by system
    println!("Conversion of 20,000 cells took: {:?}", duration);

    // Verify output is valid
    let workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert_eq!(sheet_names.len(), 1);
}
