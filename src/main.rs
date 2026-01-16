use anyhow::Result;
use clap::Parser;
use sqlite_to_xlsx::{convert, BlobHandling, ConvertOptions};
use std::path::{Path, PathBuf};

#[derive(Parser)]
#[command(name = "sqlite-to-xlsx")]
#[command(about = "Convert SQLite databases to Excel spreadsheets")]
pub struct Args {
    /// Input SQLite database file
    #[arg(required = true)]
    pub input: PathBuf,

    /// Output Excel file (default: input with .xlsx extension)
    #[arg(short, long)]
    pub output: Option<PathBuf>,

    /// Only export these tables (comma-separated)
    #[arg(short, long, value_delimiter = ',')]
    pub tables: Option<Vec<String>>,

    /// Exclude these tables (comma-separated)
    #[arg(short, long, value_delimiter = ',')]
    pub exclude: Option<Vec<String>>,

    /// BLOB handling mode
    #[arg(long, default_value = "placeholder", value_parser = parse_blob_handling)]
    pub blob_mode: BlobHandling,

    /// Don't write column headers
    #[arg(long)]
    pub no_headers: bool,

    /// Suppress all output except errors
    #[arg(long)]
    pub quiet: bool,

    /// Custom SQL query to execute and export to a named sheet
    #[arg(short = 'q', long, num_args = 1)]
    pub query: Vec<String>,

    /// Sheet name for corresponding query (must match number of --query args)
    #[arg(short = 's', long, num_args = 1)]
    pub sheet: Vec<String>,
}

/// Parse a string into a BlobHandling value
fn parse_blob_handling(s: &str) -> Result<BlobHandling, String> {
    match s.to_lowercase().as_str() {
        "placeholder" => Ok(BlobHandling::Placeholder),
        "hex" => Ok(BlobHandling::Hex),
        "base64" => Ok(BlobHandling::Base64),
        "skip" => Ok(BlobHandling::Skip),
        _ => Err(format!(
            "Invalid blob mode: {}. Valid options are: placeholder, hex, base64, skip",
            s
        )),
    }
}

fn main() {
    // 1. Parse args
    let args = Args::parse();

    // 2. Check if input file exists (for better error message)
    if !args.input.exists() {
        eprintln!("Error: Database file not found: {}\n\nMake sure the file exists and path is correct.", args.input.display());
        std::process::exit(1);
    }

    // 3. Validate query/sheet arguments match
    if args.query.len() != args.sheet.len() {
        eprintln!("Error: Number of --query arguments ({}) must match number of --sheet arguments ({})\n\nExample: sqlite-to-xlsx data.db --query 'SELECT * FROM users WHERE active=1' --sheet 'Active Users'", args.query.len(), args.sheet.len());
        std::process::exit(1);
    }

    // 4. Determine output path (default: input.xlsx)
    let output_path = if let Some(output) = args.output {
        output
    } else {
        let mut output = args.input.clone();
        output.set_extension("xlsx");
        output
    };

    // 5. Build ConvertOptions from args
    let queries: Vec<(String, String)> = args.query.into_iter()
        .zip(args.sheet)
        .collect();

    let options = ConvertOptions {
        tables: args.tables.clone(),
        exclude: args.exclude,
        blob_handling: args.blob_mode,
        write_headers: !args.no_headers,
        quiet: args.quiet,
        queries,
    };

    // 6. Call convert() with error handling
    if let Err(e) = convert(&args.input, &output_path, &options) {
        print_error(&e, &args.input, &args.tables);
        std::process::exit(1);
    }
}

fn print_error(error: &anyhow::Error, input_path: &Path, requested_tables: &Option<Vec<String>>) {
    let error_msg = error.to_string();

    // Check for specific error patterns and provide friendly messages
    if error_msg.contains("Cannot open database") {
        eprintln!("Error: Cannot open database: {}\n\nFile may not be a valid SQLite database or is locked.", input_path.display());
    } else if error_msg.contains("Cannot write to output file") {
        eprintln!("Error: {}\n\nCheck directory permissions or try a different location.", error_msg);
    } else if error_msg.contains("No tables found") {
        eprintln!("Error: No tables found in database\n\nThe database may be empty or contain only views.");
    } else if error_msg.contains("Query failed") || error_msg.contains("SQL error") {
        eprintln!("Error: {}\n\nCheck your SQL syntax and ensure referenced tables exist.", error_msg);
    } else if let Some(ref tables) = requested_tables {
        if error_msg.contains("no such table") || error_msg.contains("does not exist") {
            // Extract the first available table name from the error if available
            eprintln!("Error: Table not found: {}\n\nAvailable tables: [check database contents]", tables.join(", "));
            return;
        }
    }

    // Default error message
    eprintln!("Error: {}", error_msg);
}
