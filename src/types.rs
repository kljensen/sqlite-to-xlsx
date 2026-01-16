// Type definitions for the SQLite to XLSX converter

use rusqlite::Connection;
use anyhow::Result;

/// Information about a table in the SQLite database
pub struct TableInfo {
    pub name: String,
    pub columns: Vec<ColumnInfo>,
}

/// Information about a column in a SQLite table
pub struct ColumnInfo {
    pub name: String,
    pub sqlite_type: String,
    pub is_primary_key: bool,
}

/// Discovers all tables in the SQLite database and retrieves their schema information
///
/// # Arguments
/// * `conn` - A reference to the SQLite connection
///
/// # Returns
/// A vector of TableInfo containing table names and their column definitions
///
/// # Examples
/// ```no_run
/// use rusqlite::Connection;
/// use sqlite_to_xlsx::discover_tables;
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let conn = Connection::open("database.db")?;
/// let tables = discover_tables(&conn)?;
/// for table in tables {
///     println!("Table: {}", table.name);
///     for column in table.columns {
///         println!("  Column: {} ({})", column.name, column.sqlite_type);
///     }
/// }
/// # Ok(())
/// # }
/// ```
pub fn discover_tables(conn: &Connection) -> Result<Vec<TableInfo>> {
    let mut tables = Vec::new();

    // Query sqlite_master to get all tables (excluding sqlite_* internal tables)
    let mut stmt = conn.prepare(
        "SELECT name FROM sqlite_master
         WHERE type = 'table'
         AND name NOT LIKE 'sqlite_%'
         ORDER BY name"
    )?;

    let table_names = stmt.query_map([], |row| row.get(0))?;

    // Collect table names first to avoid statement lifetime issues
    let table_names: Vec<String> = table_names.filter_map(|name| name.ok()).collect();

    // Get column information for each table
    for table_name in table_names {
        let columns = get_table_columns(conn, &table_name)?;
        tables.push(TableInfo {
            name: table_name,
            columns,
        });
    }

    Ok(tables)
}

/// Retrieves column information for a specific table using PRAGMA table_info
///
/// # Arguments
/// * `conn` - A reference to the SQLite connection
/// * `table_name` - The name of the table to query
///
/// # Returns
/// A vector of ColumnInfo with column details
fn get_table_columns(conn: &Connection, table_name: &str) -> Result<Vec<ColumnInfo>> {
    let mut columns = Vec::new();

    // PRAGMA table_info returns columns in order with:
    // cid (column id), name, type, notnull, default_value, pk (primary key)
    let mut stmt = conn.prepare(&format!("PRAGMA table_info('{}')", table_name.replace("'", "''")))?;

    let column_rows = stmt.query_map([], |row| {
        Ok((
            row.get::<_, String>(1)?,  // name
            row.get::<_, String>(2)?,  // type
            row.get::<_, i32>(5)?,     // pk (primary key flag)
        ))
    })?;

    for column_row in column_rows {
        if let Ok((name, sqlite_type, pk)) = column_row {
            columns.push(ColumnInfo {
                name,
                sqlite_type,
                is_primary_key: pk > 0,
            });
        }
    }

    Ok(columns)
}

#[cfg(test)]
mod tests {
    use super::*;
    use rusqlite::Connection;

    #[test]
    fn test_discover_tables_empty_database() {
        let conn = Connection::open_in_memory().expect("Failed to create in-memory database");
        let tables = discover_tables(&conn).expect("Failed to discover tables");
        assert_eq!(tables.len(), 0, "Empty database should have no tables");
    }

    #[test]
    fn test_discover_tables_with_only_views() {
        let conn = Connection::open_in_memory().expect("Failed to create in-memory database");

        // Create a view but no tables
        conn.execute(
            "CREATE VIEW test_view AS SELECT 1 AS value",
            []
        ).expect("Failed to create view");

        let tables = discover_tables(&conn).expect("Failed to discover tables");
        assert_eq!(tables.len(), 0, "Database with only views should have no tables");
    }

    #[test]
    fn test_discover_tables_with_no_columns() {
        let conn = Connection::open_in_memory().expect("Failed to create in-memory database");

        // SQLite requires at least one column, so we test with a minimal table
        conn.execute(
            "CREATE TABLE minimal_table (id INTEGER)",
            []
        ).expect("Failed to create table");

        let tables = discover_tables(&conn).expect("Failed to discover tables");
        assert_eq!(tables.len(), 1, "Should find one table");
        assert_eq!(tables[0].name, "minimal_table");
        assert_eq!(tables[0].columns.len(), 1, "Should have one column");
        assert_eq!(tables[0].columns[0].name, "id");
    }

    #[test]
    fn test_discover_tables_with_schema() {
        let conn = Connection::open_in_memory().expect("Failed to create in-memory database");

        // Create a table with various column types
        conn.execute(
            "CREATE TABLE users (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                email TEXT,
                age INTEGER,
                balance REAL
            )",
            []
        ).expect("Failed to create table");

        let tables = discover_tables(&conn).expect("Failed to discover tables");
        assert_eq!(tables.len(), 1, "Should find one table");

        let table = &tables[0];
        assert_eq!(table.name, "users");
        assert_eq!(table.columns.len(), 5, "Should have 5 columns");

        // Verify column details
        assert_eq!(table.columns[0].name, "id");
        assert_eq!(table.columns[0].sqlite_type, "INTEGER");
        assert!(table.columns[0].is_primary_key, "id should be primary key");

        assert_eq!(table.columns[1].name, "name");
        assert_eq!(table.columns[1].sqlite_type, "TEXT");
        assert!(!table.columns[1].is_primary_key, "name should not be primary key");

        assert_eq!(table.columns[2].name, "email");
        assert_eq!(table.columns[2].sqlite_type, "TEXT");

        assert_eq!(table.columns[3].name, "age");
        assert_eq!(table.columns[3].sqlite_type, "INTEGER");

        assert_eq!(table.columns[4].name, "balance");
        assert_eq!(table.columns[4].sqlite_type, "REAL");
    }

    #[test]
    fn test_discover_tables_multiple_tables() {
        let conn = Connection::open_in_memory().expect("Failed to create in-memory database");

        // Create multiple tables
        conn.execute(
            "CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT)",
            []
        ).expect("Failed to create users table");

        conn.execute(
            "CREATE TABLE products (id INTEGER PRIMARY KEY, name TEXT, price REAL)",
            []
        ).expect("Failed to create products table");

        conn.execute(
            "CREATE TABLE orders (id INTEGER PRIMARY KEY, user_id INTEGER, product_id INTEGER)",
            []
        ).expect("Failed to create orders table");

        let tables = discover_tables(&conn).expect("Failed to discover tables");
        assert_eq!(tables.len(), 3, "Should find three tables");

        // Tables should be sorted alphabetically
        assert_eq!(tables[0].name, "orders");
        assert_eq!(tables[1].name, "products");
        assert_eq!(tables[2].name, "users");
    }

    #[test]
    fn test_discover_tables_excludes_sqlite_internal_tables() {
        let conn = Connection::open_in_memory().expect("Failed to create in-memory database");

        // Create a user table
        conn.execute(
            "CREATE TABLE my_table (id INTEGER PRIMARY KEY)",
            []
        ).expect("Failed to create table");

        let tables = discover_tables(&conn).expect("Failed to discover tables");

        // Should only find our custom table, not sqlite_* tables
        assert_eq!(tables.len(), 1, "Should only find custom tables");
        assert_eq!(tables[0].name, "my_table");
    }

    #[test]
    fn test_get_table_columns_with_primary_key() {
        let conn = Connection::open_in_memory().expect("Failed to create in-memory database");

        conn.execute(
            "CREATE TABLE test_table (
                id INTEGER PRIMARY KEY,
                user_id INTEGER,
                name TEXT
            )",
            []
        ).expect("Failed to create table");

        let columns = get_table_columns(&conn, "test_table").expect("Failed to get columns");
        assert_eq!(columns.len(), 3);

        assert!(columns[0].is_primary_key, "First column (id) should be primary key");
        assert!(!columns[1].is_primary_key, "Second column (user_id) should not be primary key");
        assert!(!columns[2].is_primary_key, "Third column (name) should not be primary key");
    }

    #[test]
    fn test_open_database_read_only() {
        use tempfile::TempDir;

        // Create a temporary database file
        let temp_dir = TempDir::new().expect("Failed to create temp dir");
        let db_path = temp_dir.path().join("test.db");

        {
            // Create and populate the database
            let conn = Connection::open(&db_path).expect("Failed to create database");
            conn.execute(
                "CREATE TABLE test (id INTEGER PRIMARY KEY, value TEXT)",
                []
            ).expect("Failed to create table");
            conn.execute(
                "INSERT INTO test (id, value) VALUES (1, 'test')",
                []
            ).expect("Failed to insert data");
        }

        // Open in read-only mode
        let flags = rusqlite::OpenFlags::SQLITE_OPEN_READ_ONLY
            | rusqlite::OpenFlags::SQLITE_OPEN_NO_MUTEX;
        let conn = Connection::open_with_flags(&db_path, flags)
            .expect("Failed to open database in read-only mode");

        let tables = discover_tables(&conn).expect("Failed to discover tables");
        assert_eq!(tables.len(), 1, "Should find one table");

        // Verify we can read but not write
        let result = conn.execute("INSERT INTO test (id, value) VALUES (2, 'fail')", []);
        assert!(result.is_err(), "Should not be able to write to read-only database");
    }

    #[test]
    fn test_column_info_with_various_types() {
        let conn = Connection::open_in_memory().expect("Failed to create in-memory database");

        // Create table with various SQLite types
        conn.execute(
            "CREATE TABLE type_test (
                col_text TEXT,
                col_integer INTEGER,
                col_real REAL,
                col_blob BLOB,
                col_numeric NUMERIC
            )",
            []
        ).expect("Failed to create table");

        let tables = discover_tables(&conn).expect("Failed to discover tables");
        assert_eq!(tables.len(), 1);

        let columns = &tables[0].columns;
        assert_eq!(columns.len(), 5);

        assert_eq!(columns[0].sqlite_type, "TEXT");
        assert_eq!(columns[1].sqlite_type, "INTEGER");
        assert_eq!(columns[2].sqlite_type, "REAL");
        assert_eq!(columns[3].sqlite_type, "BLOB");
        assert_eq!(columns[4].sqlite_type, "NUMERIC");
    }

    #[test]
    fn test_table_name_with_special_characters() {
        let conn = Connection::open_in_memory().expect("Failed to create in-memory database");

        // Create table with special characters in name
        conn.execute(
            "CREATE TABLE \"table-with-dash\" (id INTEGER PRIMARY KEY)",
            []
        ).expect("Failed to create table");

        let tables = discover_tables(&conn).expect("Failed to discover tables");
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "table-with-dash");
    }

    #[test]
    fn test_column_name_with_special_characters() {
        let conn = Connection::open_in_memory().expect("Failed to create in-memory database");

        // Create table with special characters in column names
        conn.execute(
            "CREATE TABLE special_cols (
                id INTEGER PRIMARY KEY,
                \"column-with-dash\" TEXT,
                \"column.with.dot\" INTEGER
            )",
            []
        ).expect("Failed to create table");

        let tables = discover_tables(&conn).expect("Failed to discover tables");
        assert_eq!(tables.len(), 1);

        let columns = &tables[0].columns;
        assert_eq!(columns.len(), 3);
        assert_eq!(columns[0].name, "id");
        assert_eq!(columns[1].name, "column-with-dash");
        assert_eq!(columns[2].name, "column.with.dot");
    }
}
