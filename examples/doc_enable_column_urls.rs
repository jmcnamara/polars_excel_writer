// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates converting strings to urls.

use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Site" => &["Rust home", "Crates.io", "Docs.rs", "Polars"],
        "Link" => &["https://www.rust-lang.org/",
                    "https://crates.io/",
                    "https://docs.rs/",
                    "https://pola.rs/"],
    )?;

    // Create a new Excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Treat a string column as a list of URLs.
    excel_writer.enable_column_urls("Link");

    // Autofit the output data.
    excel_writer.set_autofit(true);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
