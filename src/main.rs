use anyhow::Result;
use clap::Parser;
use sqlite_to_xlsx::{convert, BlobHandling, ConvertOptions};
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "sqlite2xlsx")]
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

fn main() -> Result<()> {
    // 1. Parse args
    let args = Args::parse();

    // 2. Determine output path (default: input.xlsx)
    let output_path = if let Some(output) = args.output {
        output
    } else {
        let mut output = args.input.clone();
        output.set_extension("xlsx");
        output
    };

    // 3. Build ConvertOptions from args
    let options = ConvertOptions {
        tables: args.tables,
        exclude: args.exclude,
        blob_handling: args.blob_mode,
        write_headers: !args.no_headers,
        quiet: args.quiet,
    };

    // 4. Call convert()
    convert(&args.input, &output_path, &options)?;

    Ok(())
}
