// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates handling NaN and Infinity values with custom string
//! representations.
use polars::prelude::*;

use polars_excel_writer::PolarsXlsxWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Default" => &["NAN", "INF", "-INF"],
        "Custom" => &[f64::NAN, f64::INFINITY, f64::NEG_INFINITY],
    )?;

    // Write the dataframe to an Excel file.
    let mut xlsx_writer = PolarsXlsxWriter::new();

    // Set custom values for NaN, Infinity, and -Infinity.
    xlsx_writer.set_nan_value("NaN");
    xlsx_writer.set_infinity_value("Infinity");
    xlsx_writer.set_neg_infinity_value("-Infinity");

    // Autofit the output data, for clarity.
    xlsx_writer.set_autofit(true);

    // Write the dataframe to Excel.
    xlsx_writer.write_dataframe(&df)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
