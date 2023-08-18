// excel_writer - A Polars extension to serialize dataframes to Excel xlsx files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::path::Path;

use polars::export::arrow::temporal_conversions::{
    date32_to_date, time64ns_to_time, timestamp_ms_to_datetime, timestamp_ns_to_datetime,
    timestamp_us_to_datetime,
};
use polars::prelude::*;
use rust_xlsxwriter::{Format, Table, Workbook, Worksheet, XlsxError};

/// TODO
pub struct PolarsXlsxWriter {
    pub(crate) workbook: Workbook,
    pub(crate) options: WriterOptions,
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
        let mut workbook = Workbook::new();
        workbook.add_worksheet();

        PolarsXlsxWriter {
            workbook,
            options: WriterOptions::default(),
        }
    }

    /// TODO
    ///
    /// # Errors
    ///
    pub fn write_excel<P: AsRef<Path>>(&mut self, path: P) -> PolarsResult<()> {
        self.workbook.save(path)?;

        Ok(())
    }

    /// TODO
    ///
    /// # Errors
    ///
    pub fn write_dataframe(&mut self, df: &DataFrame) -> PolarsResult<()> {
        let options = self.options.clone();
        let worksheet = self.last_worksheet()?;

        Self::write_dataframe_internal(df, worksheet, 0, 0, &options)?;

        Ok(())
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
        let options = self.options.clone();
        let worksheet = self.last_worksheet()?;

        Self::write_dataframe_internal(df, worksheet, row, col, &options)?;

        Ok(())
    }

    /// TODO
    ///
    /// # Errors
    ///
    pub fn write_dataframe_to_worksheet(
        &mut self,
        df: &DataFrame,
        worksheet: &mut Worksheet,
        row: u32,
        col: u16,
    ) -> PolarsResult<()> {
        let options = self.options.clone();

        Self::write_dataframe_internal(df, worksheet, row, col, &options)?;

        Ok(())
    }

    /// TODO
    pub fn has_header(&mut self, has_header: bool) -> &mut PolarsXlsxWriter {
        self.options.has_header = has_header;
        self.options.table.set_header_row(has_header);

        self
    }

    /// TODO
    pub fn use_autofit(&mut self, use_autofit: bool) -> &mut PolarsXlsxWriter {
        self.options.use_autofit = use_autofit;
        self
    }

    // -----------------------------------------------------------------------
    // Internal functions/methods.
    // -----------------------------------------------------------------------

    fn last_worksheet(&mut self) -> PolarsResult<&mut Worksheet> {
        let mut last_index = self.workbook.worksheets().len();

        // Add a worksheet if there isn't one already.
        if last_index == 0 {
            self.workbook.add_worksheet();
        } else {
            last_index -= 1;
        }

        let worksheet = self.workbook.worksheet_from_index(last_index)?;

        Ok(worksheet)
    }

    // TODO
    #[allow(clippy::too_many_lines)]
    fn write_dataframe_internal(
        df: &DataFrame,
        worksheet: &mut Worksheet,
        row_offset: u32,
        col_offset: u16,
        options: &WriterOptions,
    ) -> Result<(), XlsxError> {
        let header_offset = u32::from(options.has_header);

        // Iterate through the dataframe column by column.
        for (col_num, column) in df.get_columns().iter().enumerate() {
            let col_num = col_offset + col_num as u16;

            // Store the column names for use as table headers.
            if options.has_header {
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
                            &options.float_format,
                        )?;
                    }
                    AnyValue::Float64(value) => {
                        worksheet.write_number_with_format(
                            row_num,
                            col_num,
                            value,
                            &options.float_format,
                        )?;
                    }
                    AnyValue::Utf8(value) => {
                        worksheet.write_string(row_num, col_num, value)?;
                    }
                    AnyValue::Boolean(value) => {
                        worksheet.write_boolean(row_num, col_num, value)?;
                    }
                    AnyValue::Null => {
                        if let Some(null_string) = &options.null_string {
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
                            &options.datetime_format,
                        )?;
                        worksheet.set_column_width(col_num, 18)?;
                    }
                    AnyValue::Date(value) => {
                        let date = date32_to_date(value);
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            &date,
                            &options.date_format,
                        )?;
                        worksheet.set_column_width(col_num, 10)?;
                    }
                    AnyValue::Time(value) => {
                        let time = time64ns_to_time(value);
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            &time,
                            &options.time_format,
                        )?;
                    }
                    _ => {
                        eprintln!(
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
        if !options.has_header {
            max_row -= 1;
        }

        // Add the table to the worksheet.
        worksheet.add_table(
            row_offset,
            col_offset,
            row_offset + max_row as u32,
            col_offset + max_col as u16 - 1,
            &options.table,
        )?;

        // Autofit the columns.
        if options.use_autofit {
            worksheet.autofit();
        }

        Ok(())
    }

    // -----------------------------------------------------------------------
    // TODO
    // -----------------------------------------------------------------------

    // TODO
    pub(crate) fn write_to_buffer(&mut self, df: &DataFrame) -> PolarsResult<Vec<u8>> {
        let options = self.options.clone();
        let worksheet = self.last_worksheet()?;

        Self::write_dataframe_internal(df, worksheet, 0, 0, &options)?;

        let buf = self.workbook.save_to_buffer()?;

        Ok(buf)
    }
}

// -----------------------------------------------------------------------
// TODO
// -----------------------------------------------------------------------

/// TODO
#[derive(Clone)]
pub(crate) struct WriterOptions {
    pub(crate) has_header: bool,
    pub(crate) use_autofit: bool,
    pub(crate) date_format: Format,
    pub(crate) time_format: Format,
    pub(crate) float_format: Format,
    pub(crate) datetime_format: Format,
    pub(crate) null_string: Option<String>,
    pub(crate) table: Table,
}

impl Default for WriterOptions {
    fn default() -> Self {
        Self::new()
    }
}

// TODO
impl WriterOptions {
    // TODO
    fn new() -> WriterOptions {
        WriterOptions {
            has_header: true,
            use_autofit: false,
            time_format: "hh:mm:ss;@".into(),
            date_format: "yyyy\\-mm\\-dd;@".into(),
            datetime_format: "yyyy\\-mm\\-dd\\ hh:mm:ss".into(),
            null_string: None,
            float_format: Format::default(),
            table: Table::new(),
        }
    }
}
