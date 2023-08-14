// excel_writer - A Polars extension to serialize dataframes to Excel xlsx files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023, John McNamara, jmcnamara@cpan.org

use std::io::Write;

use polars::export::arrow::temporal_conversions::{
    date32_to_date, time64ns_to_time, timestamp_ms_to_datetime, timestamp_ns_to_datetime,
    timestamp_us_to_datetime,
};
use polars::prelude::*;
use rust_xlsxwriter::{Format, Table, TableColumn, Workbook, XlsxError};

use crate::ExcelWriter;

pub struct PolarsXlsxWriter {
    has_header: bool,
    date_format: Format,
    time_format: Format,
    float_format: Format,
    datetime_format: Format,
    null_string: Option<String>,
}

impl Default for PolarsXlsxWriter {
    fn default() -> Self {
        Self::new()
    }
}

impl PolarsXlsxWriter {
    pub fn new() -> PolarsXlsxWriter {
        let float_format = Format::default();
        let date_format = Format::new().set_num_format("yyyy\\-mm\\-dd;@");
        let time_format = Format::new().set_num_format("hh:mm:ss;@");
        let datetime_format = Format::new().set_num_format("yyyy\\-mm\\-dd\\ hh:mm:ss");

        PolarsXlsxWriter {
            has_header: true,
            date_format,
            time_format,
            null_string: None,
            float_format,
            datetime_format,
        }
    }

    pub(crate) fn new_from_excel_writer<W: Write>(
        excel_writer: &ExcelWriter<W>,
    ) -> PolarsXlsxWriter {
        let mut xlsx_writer = PolarsXlsxWriter {
            has_header: excel_writer.has_header,
            ..Default::default()
        };

        if !excel_writer.float_format.is_empty() {
            xlsx_writer.float_format = Format::new().set_num_format(&excel_writer.float_format);
        }

        if !excel_writer.time_format.is_empty() {
            xlsx_writer.time_format = Format::new().set_num_format(&excel_writer.time_format);
        }

        if !excel_writer.date_format.is_empty() {
            xlsx_writer.date_format = Format::new().set_num_format(&excel_writer.date_format);
        }

        if !excel_writer.datetime_format.is_empty() {
            xlsx_writer.datetime_format =
                Format::new().set_num_format(&excel_writer.datetime_format);
        }

        if !excel_writer.null_string.is_empty() {
            xlsx_writer.null_string = Some(excel_writer.null_string.clone());
        }

        xlsx_writer
    }

    /// TODO
    ///
    /// # Errors
    ///
    pub fn write_buffer(&self, df: &DataFrame) -> Result<Vec<u8>, XlsxError> {
        let mut workbook = self.create_xlsx_file(df)?;

        let buf = workbook.save_to_buffer().unwrap();

        Ok(buf)
    }

    // TODO
    fn create_xlsx_file(&self, df: &DataFrame) -> Result<Workbook, XlsxError> {
        // Create a new Excel file object.
        let mut workbook = Workbook::new();
        let mut headers = vec![];
        let row_offset = u32::from(self.has_header);

        // Add a worksheet to the workbook.
        let worksheet = workbook.add_worksheet();

        // Iterate through the dataframe column by column.
        for (col_num, column) in df.get_columns().iter().enumerate() {
            let col_num = col_num as u16;

            // Store the column names for use as table headers.
            headers.push(column.name().to_string());

            // Write the row data for each column/type.
            for (row_num, data) in column.iter().enumerate() {
                let row_num = row_offset + row_num as u32;

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
                        worksheet.write_number_with_format(
                            row_num,
                            col_num,
                            value,
                            &self.float_format,
                        )?;
                    }
                    AnyValue::Float64(value) => {
                        worksheet.write_number_with_format(
                            row_num,
                            col_num,
                            value,
                            &self.float_format,
                        )?;
                    }
                    AnyValue::Utf8(value) => {
                        worksheet.write_string(row_num, col_num, value)?;
                    }
                    AnyValue::Boolean(value) => {
                        worksheet.write_boolean(row_num, col_num, value)?;
                    }
                    AnyValue::Null => {
                        if let Some(null_string) = &self.null_string {
                            worksheet.write_string(row_num, col_num, null_string)?;
                        }
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
                            &self.datetime_format,
                        )?;
                        worksheet.set_column_width(col_num, 18)?;
                    }
                    AnyValue::Date(value) => {
                        let date = date32_to_date(value);
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            &date,
                            &self.date_format,
                        )?;
                        worksheet.set_column_width(col_num, 10)?;
                    }
                    AnyValue::Time(value) => {
                        let time = time64ns_to_time(value);
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            &time,
                            &self.time_format,
                        )?;
                    }
                    _ => {
                        println!(
                            "WARNING: AnyValue data type '{}' is not supported by Excel",
                            data.dtype()
                        );
                        break;
                    }
                }
            }
        }

        // Create a table for the dataframe range.
        let (mut max_row, max_col) = df.shape();
        let mut table = Table::new();
        if self.has_header {
            let columns: Vec<TableColumn> = headers
                .into_iter()
                .map(|x| TableColumn::new().set_header(x))
                .collect();
            table.set_columns(&columns);
        } else {
            max_row -= 1;
            table.set_header_row(false);
        }

        // Add the table to the worksheet.
        worksheet.add_table(0, 0, max_row as u32, max_col as u16 - 1, &table)?;

        // Autofit the columns.
        //worksheet.autofit();

        Ok(workbook)
    }
}
