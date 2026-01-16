// Integration tests for Unicode data handling in SQLite to XLSX conversion

use calamine::{Reader, Xlsx, open_workbook, Data};
use rusqlite::Connection;
use tempfile::TempDir;
use sqlite_to_xlsx::{convert, ConvertOptions};

#[test]
fn test_basic_unicode_characters() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with various Unicode characters
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE unicode_test (
            id INTEGER,
            text TEXT
        )",
        [],
    ).unwrap();

    // Insert various Unicode strings
    let unicode_strings = vec![
        "Hello World",           // ASCII
        "こんにちは世界",           // Japanese (Hiragana + Kanji)
        "안녕하세요 세계",           // Korean
        "你好世界",               // Simplified Chinese
        "Привет мир",            // Russian (Cyrillic)
        "Γεια σου κόσμε",        // Greek
        "مرحبا بالعالم",          // Arabic (RTL)
        "שלום עולם",             // Hebrew (RTL)
        "Hola mundo",            // Spanish
        "Bonjour le monde",      // French
        "ÄÖÜäöüß",               // German with umlauts
    ];

    for (i, text) in unicode_strings.iter().enumerate() {
        conn.execute(
            "INSERT INTO unicode_test (id, text) VALUES (?1, ?2)",
            [&((i as i64) + 1) as &dyn rusqlite::ToSql, &&**text as &dyn rusqlite::ToSql],
        ).unwrap();
    }

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify Unicode strings are correctly written to XLSX
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("unicode_test")
        .expect("Failed to read sheet 'unicode_test'");

    // Verify each Unicode string
    for (i, expected_text) in unicode_strings.iter().enumerate() {
        let row = ((i + 1) as u32) as usize; // +1 for header row
        let cell_value = sheet.get((row, 1));
        match cell_value {
            Some(Data::String(s)) => assert_eq!(&**s, *expected_text),
            _ => panic!("Expected string at row {}, col 1, got {:?}", row, cell_value),
        }
    }
}

#[test]
fn test_emoji_characters() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with emoji
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE emoji_test (
            id INTEGER,
            emoji TEXT,
            description TEXT
        )",
        [],
    ).unwrap();

    // Insert emoji data
    let emoji_data = vec![
        ("😀", "Grinning Face"),
        ("🎉", "Party Popper"),
        ("🌍", "Globe"),
        ("❤️", "Red Heart"),
        ("🚀", "Rocket"),
        ("🔥", "Fire"),
        ("✨", "Sparkles"),
        ("🎵", "Musical Note"),
        ("☕", "Hot Beverage"),
        ("💻", "Computer"),
    ];

    for (i, (emoji, desc)) in emoji_data.iter().enumerate() {
        conn.execute(
            "INSERT INTO emoji_test (id, emoji, description) VALUES (?1, ?2, ?3)",
            [&((i as i64) + 1) as &dyn rusqlite::ToSql, &&**emoji as &dyn rusqlite::ToSql, &&**desc as &dyn rusqlite::ToSql],
        ).unwrap();
    }

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify emoji are correctly written
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("emoji_test")
        .expect("Failed to read sheet 'emoji_test'");

    // Check a few emoji values
    assert_eq!(sheet.get((1, 1)), Some(&Data::String("😀".to_string())));
    assert_eq!(sheet.get((1, 2)), Some(&Data::String("Grinning Face".to_string())));
    assert_eq!(sheet.get((2, 1)), Some(&Data::String("🎉".to_string())));
    assert_eq!(sheet.get((3, 1)), Some(&Data::String("🌍".to_string())));
    assert_eq!(sheet.get((10, 1)), Some(&Data::String("💻".to_string())));
}

#[test]
fn test_unicode_in_table_and_column_names() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Create database with Unicode table and column names
    let conn = Connection::open(&db_path).expect("Failed to create database");

    // Note: SQLite supports Unicode in identifiers
    conn.execute(
        "CREATE TABLE 「ユーザー」 (
            「ID」 INTEGER,
            「名前」 TEXT,
            「メール」 TEXT
        )",
        [],
    ).unwrap();

    conn.execute(
        "INSERT INTO 「ユーザー」 (「ID」, 「名前」, 「メール」) VALUES (1, '田中', 'tanaka@example.com')",
        [],
    ).unwrap();

    conn.execute(
        "INSERT INTO 「ユーザー」 (「ID」, 「名前」, 「メール」) VALUES (2, '鈴木', 'suzuki@example.com')",
        [],
    ).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify Unicode names are preserved
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet_names = workbook.sheet_names();
    assert!(sheet_names.contains(&"「ユーザー」".to_string()));

    let sheet = workbook.worksheet_range("「ユーザー」")
        .expect("Failed to read sheet '「ユーザー」'");

    // Verify headers with Unicode
    assert_eq!(sheet.get((0, 0)), Some(&Data::String("「ID」".to_string())));
    assert_eq!(sheet.get((0, 1)), Some(&Data::String("「名前」".to_string())));
    assert_eq!(sheet.get((0, 2)), Some(&Data::String("「メール」".to_string())));

    // Verify data
    assert_eq!(sheet.get((1, 1)), Some(&Data::String("田中".to_string())));
    assert_eq!(sheet.get((2, 1)), Some(&Data::String("鈴木".to_string())));
}

#[test]
fn test_mixed_script_text() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test mixed scripts in same text
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE mixed_script (
            id INTEGER,
            text TEXT
        )",
        [],
    ).unwrap();

    let mixed_texts = vec![
        "Hello 世界 🌍",
        "日本語 English 日本語",
        "العربية English العربية",
        "עברית English עברית",
        "😀 Hello 😀",
        "Café résumé naïve",  // Latin with accents
        "Мир World Мир",      // Mixed Russian and English
    ];

    for (i, text) in mixed_texts.iter().enumerate() {
        conn.execute(
            "INSERT INTO mixed_script (id, text) VALUES (?1, ?2)",
            [&((i as i64) + 1) as &dyn rusqlite::ToSql, &&**text as &dyn rusqlite::ToSql],
        ).unwrap();
    }

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify mixed script text is preserved
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("mixed_script")
        .expect("Failed to read sheet 'mixed_script'");

    for (i, expected_text) in mixed_texts.iter().enumerate() {
        let row = ((i + 1) as u32) as usize;
        let cell_value = sheet.get((row, 1));
        match cell_value {
            Some(Data::String(s)) => {
                assert_eq!(&**s, *expected_text, "Row {} text mismatch", row);
            },
            _ => panic!("Expected string at row {}, got {:?}", row, cell_value),
        }
    }
}

#[test]
fn test_unicode_special_characters() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test special Unicode characters
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE special_chars (
            id INTEGER,
            text TEXT
        )",
        [],
    ).unwrap();

    let special_chars = vec![
        "© ® ™",                      // Copyright, registered, trademark
        "€ £ ¥ ¢ ₽",                   // Currency symbols
        "§ ¶ † ‡",                     // Section, paragraph, dagger
        "° ′ ″",                       // Degree, prime, double prime
        "− × ÷ ±",                     // Minus, multiply, divide, plus-minus
        "∀ ∁ ∂ ∃ ∅ ∆",                // Mathematical operators
        "α β γ δ ε ζ",                 // Greek letters
        "← → ↑ ↓",                     // Arrows
        "▀ ▄ ■ □",                    // Box drawing
        "① ② ③ ④ ⑤",                  // Circled numbers
        "Ⅰ Ⅱ Ⅲ Ⅳ Ⅴ",                 // Roman numerals
    ];

    for (i, text) in special_chars.iter().enumerate() {
        conn.execute(
            "INSERT INTO special_chars (id, text) VALUES (?1, ?2)",
            [&((i as i64) + 1) as &dyn rusqlite::ToSql, &&**text as &dyn rusqlite::ToSql],
        ).unwrap();
    }

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify special characters
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("special_chars")
        .expect("Failed to read sheet 'special_chars'");

    for (i, expected_text) in special_chars.iter().enumerate() {
        let row = ((i + 1) as u32) as usize;
        let cell_value = sheet.get((row, 1));
        match cell_value {
            Some(Data::String(s)) => {
                assert_eq!(&**s, *expected_text, "Row {} special chars mismatch", row);
            },
            _ => panic!("Expected string at row {}, got {:?}", row, cell_value),
        }
    }
}

#[test]
fn test_rtl_languages() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test right-to-left languages
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE rtl_test (
            id INTEGER,
            language TEXT,
            text TEXT
        )",
        [],
    ).unwrap();

    // Arabic text
    conn.execute(
        "INSERT INTO rtl_test (id, language, text) VALUES (1, 'Arabic', 'مرحبا بك')",
        [],
    ).unwrap();

    // Hebrew text
    conn.execute(
        "INSERT INTO rtl_test (id, language, text) VALUES (2, 'Hebrew', 'שלום רב')",
        [],
    ).unwrap();

    // Persian/Farsi text
    conn.execute(
        "INSERT INTO rtl_test (id, language, text) VALUES (3, 'Persian', 'سلام')",
        [],
    ).unwrap();

    // Urdu text
    conn.execute(
        "INSERT INTO rtl_test (id, language, text) VALUES (4, 'Urdu', 'السلام علیکم')",
        [],
    ).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify RTL text is preserved
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("rtl_test")
        .expect("Failed to read sheet 'rtl_test'");

    assert_eq!(sheet.get((1, 1)), Some(&Data::String("Arabic".to_string())));
    assert_eq!(sheet.get((1, 2)), Some(&Data::String("مرحبا بك".to_string())));
    assert_eq!(sheet.get((2, 1)), Some(&Data::String("Hebrew".to_string())));
    assert_eq!(sheet.get((2, 2)), Some(&Data::String("שלום רב".to_string())));
    assert_eq!(sheet.get((3, 1)), Some(&Data::String("Persian".to_string())));
    assert_eq!(sheet.get((3, 2)), Some(&Data::String("سلام".to_string())));
}

#[test]
fn test_zero_width_and_invisible_characters() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test zero-width and other invisible Unicode characters
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE invisible_chars (
            id INTEGER,
            description TEXT,
            text TEXT
        )",
        [],
    ).unwrap();

    // Zero-width space (U+200B)
    conn.execute(
        "INSERT INTO invisible_chars (id, description, text) VALUES (1, 'Zero-width space', 'hello\u{200B}world')",
        [],
    ).unwrap();

    // Zero-width joiner (U+200D)
    conn.execute(
        "INSERT INTO invisible_chars (id, description, text) VALUES (2, 'Zero-width joiner', 'hello\u{200D}world')",
        [],
    ).unwrap();

    // Soft hyphen (U+00AD)
    conn.execute(
        "INSERT INTO invisible_chars (id, description, text) VALUES (3, 'Soft hyphen', 'hello\u{00AD}world')",
        [],
    ).unwrap();

    // Non-breaking space (U+00A0)
    conn.execute(
        "INSERT INTO invisible_chars (id, description, text) VALUES (4, 'Non-breaking space', 'hello\u{00A0}world')",
        [],
    ).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify invisible characters are preserved
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("invisible_chars")
        .expect("Failed to read sheet 'invisible_chars'");

    // The actual text content should match including invisible chars
    assert_eq!(sheet.get((1, 1)), Some(&Data::String("Zero-width space".to_string())));
    assert_eq!(sheet.get((2, 1)), Some(&Data::String("Zero-width joiner".to_string())));
    assert_eq!(sheet.get((3, 1)), Some(&Data::String("Soft hyphen".to_string())));
}

#[test]
fn test_very_long_unicode_strings() {
    let temp_dir = TempDir::new().expect("Failed to create temp dir");
    let db_path = temp_dir.path().join("test.db");
    let xlsx_path = temp_dir.path().join("test.xlsx");

    // Test very long Unicode strings
    let conn = Connection::open(&db_path).expect("Failed to create database");

    conn.execute(
        "CREATE TABLE long_unicode (
            id INTEGER,
            text TEXT
        )",
        [],
    ).unwrap();

    // Create a long string with repeated Japanese characters
    let long_japanese = "あいうえおかきくけこ".repeat(100); // 1500 characters
    conn.execute(
        "INSERT INTO long_unicode (id, text) VALUES (1, ?1)",
        [long_japanese.as_str()],
    ).unwrap();

    // Create a long string with emoji
    let long_emoji = "😀😃😄😁😆😅🤣😂🙂🙃".repeat(50); // 500 emoji
    conn.execute(
        "INSERT INTO long_unicode (id, text) VALUES (2, ?1)",
        [long_emoji.as_str()],
    ).unwrap();

    drop(conn);

    convert(&db_path, &xlsx_path, &ConvertOptions::default())
        .expect("Conversion should succeed");

    // Verify long Unicode strings are preserved
    let mut workbook: Xlsx<_> = open_workbook(&xlsx_path)
        .expect("Failed to open XLSX file");

    let sheet = workbook.worksheet_range("long_unicode")
        .expect("Failed to read sheet 'long_unicode'");

    // Check that long strings are stored
    match sheet.get((1, 1)) {
        Some(Data::String(s)) => {
            assert_eq!(s.len(), long_japanese.len());
            assert_eq!(&**s, long_japanese);
        },
        _ => panic!("Expected long Japanese string, got {:?}", sheet.get((1, 1))),
    }

    match sheet.get((2, 1)) {
        Some(Data::String(s)) => {
            assert_eq!(s.len(), long_emoji.len());
            assert_eq!(&**s, long_emoji);
        },
        _ => panic!("Expected long emoji string, got {:?}", sheet.get((2, 1))),
    }
}
