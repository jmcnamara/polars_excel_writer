// excel_writer - A Polars extension to serialize dataframes to Excel xlsx files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::io::Write;
use std::path::Path;

use polars::export::arrow::temporal_conversions::{
    date32_to_date, time64ns_to_time, timestamp_ms_to_datetime, timestamp_ns_to_datetime,
    timestamp_us_to_datetime,
};
use polars::prelude::*;
use rust_xlsxwriter::{Format, Table, Workbook, XlsxError};

use crate::ExcelWriter;

/// TODO
pub struct PolarsXlsxWriter {
    has_header: bool,
    use_autofit: bool,
    date_format: Format,
    time_format: Format,
    float_format: Format,
    datetime_format: Format,
    null_string: Option<String>,
    workbook: Workbook,
    table: Table,
}

impl Default for PolarsXlsxWriter {
    fn default() -> Self {
        Self::new()
    }
}

/// TODO
impl PolarsXlsxWriter {
    /// TODO
    pub fn new() -> PolarsXlsxWriter {
        let float_format = Format::default();
        let date_format = Format::new().set_num_format("yyyy\\-mm\\-dd;@");
        let time_format = Format::new().set_num_format("hh:mm:ss;@");
        let datetime_format = Format::new().set_num_format("yyyy\\-mm\\-dd\\ hh:mm:ss");

        let mut workbook = Workbook::new();
        workbook.add_worksheet();

        PolarsXlsxWriter {
            has_header: true,
            use_autofit: false,
            date_format,
            time_format,
            null_string: None,
            float_format,
            datetime_format,
            workbook,
            table: Table::new(),
        }
    }

    /// TODO
    ///
    /// # Errors
    ///
    pub fn write_excel<P: AsRef<Path>>(&mut self, path: P) -> PolarsResult<()> {
        self.workbook.save(path).unwrap();

        Ok(())
    }

    /// TODO
    ///
    /// # Errors
    ///
    pub fn write_dataframe(&mut self, df: &DataFrame) -> PolarsResult<()> {
        self.write_dataframe_internal(df, 0, 0)
            .map_err(|err| polars_err!(ComputeError: "rust_xlsxwriter error: '{}'", err))
    }

    /// TODO
    ///
    /// # Errors
    ///
    pub fn write_dataframe_to_cell(
        &mut self,
        df: &DataFrame,
        row: u32,
        col: u16,
    ) -> PolarsResult<()> {
        self.write_dataframe_internal(df, row, col)
            .map_err(|err| polars_err!(ComputeError: "rust_xlsxwriter error: '{}'", err))
    }

    /// TODO
    pub fn has_header(&mut self, has_header: bool) -> &mut PolarsXlsxWriter {
        self.has_header = has_header;
        self.table.set_header_row(has_header);
        if !has_header {
            self.table.set_autofilter(false);
        }

        self
    }

    /// TODO
    pub fn use_autofit(&mut self, use_autofit: bool) -> &mut PolarsXlsxWriter {
        self.use_autofit = use_autofit;
        self
    }

    // -----------------------------------------------------------------------
    // Internal functions/methods.
    // -----------------------------------------------------------------------

    // TODO
    #[allow(clippy::too_many_lines)]
    fn write_dataframe_internal(
        &mut self,
        df: &DataFrame,
        row_offset: u32,
        col_offset: u16,
    ) -> Result<(), XlsxError> {
        let header_offset = u32::from(self.has_header);

        // Add a worksheet to the workbook.
        let last_index = self.workbook.worksheets().len() - 1;
        let worksheet = self.workbook.worksheet_from_index(last_index)?;

        // Iterate through the dataframe column by column.
        for (col_num, column) in df.get_columns().iter().enumerate() {
            let col_num = col_offset + col_num as u16;

            // Store the column names for use as table headers.
            if self.has_header {
                worksheet.write(row_offset, col_num, column.name())?;
            }

            // Write the row data for each column/type.
            for (row_num, data) in column.iter().enumerate() {
                let row_num = header_offset + row_offset + row_num as u32;

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
                            "WARNING: Polars AnyValue data type '{}' is not supported by Excel",
                            data.dtype()
                        );
                        break;
                    }
                }
            }
        }

        // Create a table for the dataframe range.
        let (mut max_row, max_col) = df.shape();
        if !self.has_header {
            max_row -= 1;
        }

        // Add the table to the worksheet.
        worksheet.add_table(
            row_offset,
            col_offset,
            row_offset + max_row as u32,
            col_offset + max_col as u16 - 1,
            &self.table,
        )?;

        // Autofit the columns.
        if self.use_autofit {
            worksheet.autofit();
        }

        Ok(())
    }

    // -----------------------------------------------------------------------
    // TODO
    // -----------------------------------------------------------------------

    // TODO
    pub(crate) fn new_from_excel_writer<W: Write>(
        excel_writer: &ExcelWriter<W>,
    ) -> PolarsXlsxWriter {
        let mut xlsx_writer = PolarsXlsxWriter {
            has_header: excel_writer.has_header,
            use_autofit: excel_writer.has_autofit,
            ..Default::default()
        };

        if !excel_writer.has_header {
            xlsx_writer.table.set_header_row(false);
            xlsx_writer.table.set_autofilter(false);
        }

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

    // TODO
    pub(crate) fn write_to_buffer(&mut self, df: &DataFrame) -> Result<Vec<u8>, XlsxError> {
        self.write_dataframe_internal(df, 0, 0)?;

        self.workbook.save_to_buffer()
    }
}
