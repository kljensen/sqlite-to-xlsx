# sqlite2xlsx

[![CI](https://img.shields.io/github/actions/workflow/status/kljensen/sqlite2xlsx/ci.yml?branch=main&style=for-the-badge&logo=github)](https://github.com/kljensen/sqlite2xlsx/actions/workflows/ci.yml)
[![Crates.io](https://img.shields.io/crates/v/sqlite2xlsx?style=for-the-badge&logo=rust)](https://crates.io/crates/sqlite2xlsx)
[![License](https://img.shields.io/badge/license-MIT-blue.svg?style=for-the-badge)](LICENSE)

A fast, ergonomic command-line tool to convert SQLite databases into Excel spreadsheets.

## Features

- 📊 **Simple conversion** - Turn any SQLite database into an Excel file with one command
- 🔍 **Custom queries** - Export SQL query results to named sheets
- 🎯 **Table filtering** - Include or exclude specific tables
- 🖼️ **BLOB handling** - Multiple options for binary data (placeholder, hex, base64, skip)
- 📋 **Flexible output** - Optional headers, quiet mode for scripting
- ⚡ **Zero dependencies** - Single binary, no runtime dependencies

## Installation

### From crates.io (recommended)

```bash
cargo install sqlite2xlsx
```

### From source

```bash
git clone https://github.com/kljensen/sqlite2xlsx.git
cd sqlite2xlsx
cargo install --path .
```

### Pre-built binaries

Download from the [Releases](https://github.com/kljensen/sqlite2xlsx/releases) page.

## Quick Start

Basic usage - convert an entire database:

```bash
sqlite2xlsx mydata.db
```

This creates `mydata.xlsx` with all tables as sheets.

Specify output file:

```bash
sqlite2xlsx mydata.db -o report.xlsx
```

## Usage Examples

### Export specific tables

```bash
sqlite2xlsx mydata.db --tables users,orders,products
```

### Exclude certain tables

```bash
sqlite2xlsx mydata.db --exclude sqlite_sequence,temp_logs
```

### Custom SQL queries

```bash
sqlite2xlsx mydata.db -q "SELECT * FROM users WHERE active=1" -s "Active Users" \
                       -q "SELECT COUNT(*) as total FROM orders" -s "Order Count"
```

### Handle BLOB data

```bash
# Show BLOBs as base64-encoded strings
sqlite2xlsx mydata.db --blob-mode base64

# Skip BLOB columns entirely
sqlite2xlsx mydata.db --blob-mode skip

# Show BLOBs as hexadecimal (default: placeholder text)
sqlite2xlsx mydata.db --blob-mode hex
```

### Quiet mode for scripts

```bash
sqlite2xlsx mydata.db --quiet
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

## License

MIT
