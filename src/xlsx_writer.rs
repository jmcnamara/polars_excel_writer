// excel_writer - A Polars extension to serialize dataframes to Excel xlsx files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023, John McNamara, jmcnamara@cpan.org

use polars::export::arrow::temporal_conversions::{
    date32_to_date, time64ns_to_time, timestamp_ms_to_datetime, timestamp_ns_to_datetime,
    timestamp_us_to_datetime,
};
use polars::prelude::*;
use rust_xlsxwriter::{Format, Table, TableColumn, Workbook, XlsxError};

pub struct PolarsXlsxWriter {}


impl Default for PolarsXlsxWriter {
   fn default() -> Self {
       Self::new()
   }
}


impl PolarsXlsxWriter {
    pub fn new() -> PolarsXlsxWriter {
        PolarsXlsxWriter {}
    }

    pub fn write_buffer(&self, df: &DataFrame) -> Result<Vec<u8>, XlsxError> {
        // Create a new Excel file object.
        let mut workbook = Workbook::new();
        let mut headers = vec![];

        // Create some formats for the dataframe.
        let datetime_format = Format::new().set_num_format("yyyy\\-mm\\-dd\\ hh:mm:ss");
        let date_format = Format::new().set_num_format("yyyy\\-mm\\-dd;@");
        let time_format = Format::new().set_num_format("hh:mm:ss;@");

        // Add a worksheet to the workbook.
        let worksheet = workbook.add_worksheet();

        // Iterate through the dataframe column by column.
        for (col_num, column) in df.get_columns().iter().enumerate() {
            let col_num = col_num as u16;

            // Store the column names for use as table headers.
            headers.push(column.name().to_string());

            // Write the row data for each column/type.
            for (row_num, data) in column.iter().enumerate() {
                let row_num = 1 + row_num as u32;

                // Map the Polars Series AnyValue types to Excel/rust_xlsxwriter
                // types.
                match data {
                    AnyValue::Int8(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::UInt8(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::Int16(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::UInt16(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::Int32(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::UInt32(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::Float32(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::Float64(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::Utf8(value) => {
                        worksheet.write_string(row_num, col_num, value)?;
                    }
                    AnyValue::Boolean(value) => {
                        worksheet.write_boolean(row_num, col_num, value)?;
                    }
                    AnyValue::Null => {
                        // Treat Null as blank for now.
                    }
                    AnyValue::Datetime(value, time_units, _) => {
                        let datetime = match time_units {
                            TimeUnit::Nanoseconds => timestamp_ns_to_datetime(value),
                            TimeUnit::Microseconds => timestamp_us_to_datetime(value),
                            TimeUnit::Milliseconds => timestamp_ms_to_datetime(value),
                        };
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            &datetime,
                            &datetime_format,
                        )?;
                        worksheet.set_column_width(col_num, 18)?;
                    }
                    AnyValue::Date(value) => {
                        let date = date32_to_date(value);
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            &date,
                            &date_format,
                        )?;
                        worksheet.set_column_width(col_num, 10)?;
                    }
                    AnyValue::Time(value) => {
                        let time = time64ns_to_time(value);
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            &time,
                            &time_format,
                        )?;
                    }
                    _ => {
                        println!(
                            "WARNING: AnyValue data type '{}' is not supported",
                            data.dtype()
                        );
                        break;
                    }
                }
            }
        }

        // Create a table for the dataframe range.
        let (max_row, max_col) = df.shape();
        let mut table = Table::new();
        let columns: Vec<TableColumn> = headers
            .into_iter()
            .map(|x| TableColumn::new().set_header(x))
            .collect();
        table.set_columns(&columns);

        // Add the table to the worksheet.
        worksheet.add_table(0, 0, max_row as u32, max_col as u16 - 1, &table)?;

        // Autofit the columns.
        worksheet.autofit();

        // Save the file to disk.
        let buf = workbook.save_to_buffer().unwrap();

        Ok(buf)
    }
}
