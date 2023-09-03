// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of using `polars_excel_writer` in conjunction with
//! `rust_xlsxwriter` to write a Polars dataframe to a worksheet and then add a
//! chart to plot the data.

use polars::prelude::*;
use polars_excel_writer::PolarsXlsxWriter;
use rust_xlsxwriter::{Chart, ChartType, Workbook};

fn main() -> PolarsResult<()> {
    // Create a sample dataframe using `Polars`
    let df: DataFrame = df!(
        "Data" => &[10, 20, 15, 25, 30, 20],
    )?;

    // Get some dataframe dimensions that we will use for the chart range.
    let row_min = 1; // Skip the header row.
    let row_max = df.height() as u32;

    // Create a new workbook and worksheet using `rust_xlsxwriter`.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write the dataframe to the worksheet using `PolarsXlsxWriter`.
    let mut xlsx_writer = PolarsXlsxWriter::new();
    xlsx_writer.write_dataframe_to_worksheet(&df, worksheet, 0, 0)?;

    // Move back to `rust_xlsxwriter` to create a new chart and have it plot the
    // range of the dataframe in the worksheet.
    let mut chart = Chart::new(ChartType::Line);
    chart
        .add_series()
        .set_values(("Sheet1", row_min, 0, row_max, 0));

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file to disk.
    workbook.save("chart.xlsx")?;

    Ok(())
}
