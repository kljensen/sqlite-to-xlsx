use anyhow::Result;

/// Generates sheet names for the XLSX file
pub fn generate_sheet_name() -> Result<String> {
    Ok(String::from("Sheet1"))
}
