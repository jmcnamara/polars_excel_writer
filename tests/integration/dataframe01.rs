// Test case that compares a file generated by polars_excel_writer with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

use crate::common;

use polars::prelude::*;
use polars_excel_writer::{ExcelWriter, PolarsXlsxWriter};
use rust_xlsxwriter::{Workbook, XlsxError};

// Compare output against target Excel file using ExcelWriter.
fn create_new_xlsx_file_1(filename: &str) -> Result<(), XlsxError> {
    let mut df: DataFrame = df!(
        "Foo" => &[1, 1, 1],
        "Bar" => &[2, 2, 2],
    )?;

    let mut file = std::fs::File::create(filename)?;

    ExcelWriter::new(&mut file).finish(&mut df)?;

    Ok(())
}

// Compare output against target Excel file using PolarsXlsxWriter.
fn create_new_xlsx_file_2(filename: &str) -> Result<(), XlsxError> {
    let df: DataFrame = df!(
        "Foo" => &[1, 1, 1],
        "Bar" => &[2, 2, 2],
    )?;

    let mut xlsx_writer = PolarsXlsxWriter::new();
    xlsx_writer.write_dataframe(&df)?;
    xlsx_writer.save(filename)?;

    Ok(())
}

// Compare output against target Excel file using rust_xlsxwriter.
fn create_new_xlsx_file_3(filename: &str) -> Result<(), XlsxError> {
    let df: DataFrame = df!(
        "Foo" => &[1, 1, 1],
        "Bar" => &[2, 2, 2],
    )?;

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let mut xlsx_writer = PolarsXlsxWriter::new();

    xlsx_writer.write_dataframe_to_worksheet(&df, worksheet, 0, 0)?;

    workbook.save(filename)?;

    Ok(())
}

// Check CSV input which should default to u64.
fn create_new_xlsx_file_4(filename: &str) -> Result<(), XlsxError> {
    let csv_string = "Foo,Bar\n1,2\n1,2\n1,2\n";
    let buffer = std::io::Cursor::new(csv_string);
    let mut df = CsvReader::new(buffer)
        .with_null_values(NullValues::AllColumnsSingle("NULL".to_string()).into())
        .finish()?;

    let mut file = std::fs::File::create(filename)?;

    ExcelWriter::new(&mut file).finish(&mut df)?;

    Ok(())
}

#[test]
fn dataframe_excelwriter01() {
    let test_runner = common::TestRunner::new()
        .set_name("dataframe01")
        .set_function(create_new_xlsx_file_1)
        .unique("1")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn dataframe_write_excel01() {
    let test_runner = common::TestRunner::new()
        .set_name("dataframe01")
        .set_function(create_new_xlsx_file_2)
        .unique("2")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn dataframe_to_worksheet01() {
    let test_runner = common::TestRunner::new()
        .set_name("dataframe01")
        .set_function(create_new_xlsx_file_3)
        .unique("3")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}

#[test]
fn dataframe_from_csv01() {
    let test_runner = common::TestRunner::new()
        .set_name("dataframe01")
        .set_function(create_new_xlsx_file_4)
        .unique("4")
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
