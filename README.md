# sqlite-to-xlsx

[![CI](https://img.shields.io/github/actions/workflow/status/kljensen/sqlite-to-xlsx/ci.yml?branch=main&style=for-the-badge&logo=github)](https://github.com/kljensen/sqlite-to-xlsx/actions/workflows/ci.yml)
[![Crates.io](https://img.shields.io/crates/v/sqlite-to-xlsx?style=for-the-badge&logo=rust)](https://crates.io/crates/sqlite-to-xlsx)
[![License](https://img.shields.io/badge/license-Unlicense-blue.svg?style=for-the-badge)](LICENSE)

A fast, ergonomic command-line tool to convert SQLite databases into Excel spreadsheets.

## Features

- 📊 **Simple conversion** - Turn any SQLite database into an Excel file with one command
- 🔍 **Custom queries** - Export SQL query results to named sheets
- 🎯 **Table filtering** - Include or exclude specific tables
- 🖼️ **BLOB handling** - Multiple options for binary data (placeholder, hex, base64, skip)
- 📋 **Flexible output** - Optional headers, quiet mode for scripting
- ⚡ **Statically linked** - Single binary, no runtime dependencies

## Installation

### From crates.io (recommended)

```bash
cargo install sqlite-to-xlsx
```

### From source

```bash
git clone https://github.com/kljensen/sqlite-to-xlsx.git
cd sqlite-to-xlsx
cargo install --path .
```

### Pre-built binaries

Download from the [Releases](https://github.com/kljensen/sqlite-to-xlsx/releases) page.

## Quick Start

Basic usage - convert an entire database:

```bash
sqlite-to-xlsx mydata.db
```

This creates `mydata.xlsx` with all tables as sheets.

Specify output file:

```bash
sqlite-to-xlsx mydata.db -o report.xlsx
```

## Usage Examples

### Export specific tables

```bash
sqlite-to-xlsx mydata.db --tables users,orders,products
```

### Exclude certain tables

```bash
sqlite-to-xlsx mydata.db --exclude sqlite_sequence,temp_logs
```

### Custom SQL queries

```bash
sqlite-to-xlsx mydata.db -q "SELECT * FROM users WHERE active=1" -s "Active Users" \
                       -q "SELECT COUNT(*) as total FROM orders" -s "Order Count"
```

### Handle BLOB data

```bash
# Show BLOBs as base64-encoded strings
sqlite-to-xlsx mydata.db --blob-mode base64

# Skip BLOB columns entirely
sqlite-to-xlsx mydata.db --blob-mode skip

# Show BLOBs as hexadecimal (default: placeholder text)
sqlite-to-xlsx mydata.db --blob-mode hex
```

### Quiet mode for scripts

```bash
sqlite-to-xlsx mydata.db --quiet
```

## CLI Options

| Option | Short | Description | Default |
|--------|-------|-------------|---------|
| `--output` | `-o` | Output Excel file | `<input>.xlsx` |
| `--tables` | `-t` | Only export these tables (comma-separated) | All tables |
| `--exclude` | `-e` | Exclude these tables (comma-separated) | None |
| `--blob-mode` | | BLOB handling: `placeholder`, `hex`, `base64`, `skip` | `placeholder` |
| `--no-headers` | | Don't write column headers | Headers on |
| `--quiet` | | Suppress all output except errors | Verbose |
| `--query` | `-q` | Custom SQL query for named sheet | None |
| `--sheet` | `-s` | Sheet name for corresponding query | Required with `-q` |
| `--help` | `-h` | Show help message | |
| `--version` | `-V` | Show version | |

### BLOB Modes

| Mode | Description |
|------|-------------|
| `placeholder` | Shows `<BLOB: N bytes>` text |
| `hex` | Hexadecimal representation |
| `base64` | Base64-encoded string |
| `skip` | Omits BLOB columns entirely |

## Security

The `--query` option executes arbitrary SQL against the database. While the database is opened in read-only mode (which prevents modifications), users should be aware that:

- Custom queries can read any data in the database
- Query execution time is unbounded (complex queries may run indefinitely)
- Always validate queries from untrusted sources before execution

## Caveats

- Table names must be simple identifiers (letters, numbers, underscores). Names that require SQL identifier quoting (like embedded quotes or brackets) are not supported.

## License

Unlicense
