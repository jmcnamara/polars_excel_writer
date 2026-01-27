// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates handling NaN and Infinity values with custom string
//! representations.
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Default" => &["NAN", "INF", "-INF"],
        "Custom" => &[f64::NAN, f64::INFINITY, f64::NEG_INFINITY],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set custom values for NaN, Infinity, and -Infinity.
    excel_writer.set_nan_value("NaN");
    excel_writer.set_infinity_value("Infinity");
    excel_writer.set_neg_infinity_value("-Infinity");

    // Autofit the output data, for clarity.
    excel_writer.set_autofit(true);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
