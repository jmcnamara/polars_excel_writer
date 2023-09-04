// Test case that compares a file generated by polars_excel_writer with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::common;

use polars::prelude::*;
use polars_excel_writer::{ExcelWriter, PolarsXlsxWriter};
use rust_xlsxwriter::{Table, XlsxError};

// Compare output against target Excel file using ExcelWriter.
fn create_new_xlsx_file_1(filename: &str) -> Result<(), XlsxError> {
    let mut df: DataFrame = df!(
        "Column1" => &["Foo", "Foo", "Foo"],
        "Column2" => &["Bar", "Bar", "Bar"],
    )?;

    let mut file = std::fs::File::create(filename)?;

    ExcelWriter::new(&mut file)
        .has_header(false)
        .with_autofit()
        .finish(&mut df)?;

    Ok(())
}

// Compare output against target Excel file using PolarsXlsxWriter.
fn create_new_xlsx_file_2(filename: &str) -> Result<(), XlsxError> {
    let df: DataFrame = df!(
        "Column1" => &["Foo", "Foo", "Foo"],
        "Column2" => &["Bar", "Bar", "Bar"],
    )?;

    let mut xlsx_writer = PolarsXlsxWriter::new();
    xlsx_writer.set_header(false);
    xlsx_writer.set_autofit(true);

    xlsx_writer.write_dataframe(&df)?;
    xlsx_writer.save(filename)?;

    Ok(())
}

// Compare using PolarsXlsxWriter and set_table().
fn create_new_xlsx_file_3(filename: &str) -> Result<(), XlsxError> {
    let df: DataFrame = df!(
        "Column1" => &["Foo", "Foo", "Foo"],
        "Column2" => &["Bar", "Bar", "Bar"],
    )?;

    let mut xlsx_writer = PolarsXlsxWriter::new();
    let mut table = Table::new();
    table.set_header_row(false);

    xlsx_writer.set_table(&table);
    xlsx_writer.set_autofit(true);

    xlsx_writer.write_dataframe(&df)?;
    xlsx_writer.save(filename)?;

    Ok(())
}

#[test]
fn dataframe_excelwriter02() {
    let test_runner = common::TestRunner::new()
        .set_name("dataframe02")
        .set_function(create_new_xlsx_file_1)
        .unique("1")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn dataframe_write_excel02() {
    let test_runner = common::TestRunner::new()
        .set_name("dataframe02")
        .set_function(create_new_xlsx_file_2)
        .unique("2")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn dataframe_write_excel02_3() {
    let test_runner = common::TestRunner::new()
        .set_name("dataframe02")
        .set_function(create_new_xlsx_file_3)
        .unique("3")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
