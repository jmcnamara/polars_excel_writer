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

    // TODO
    pub(crate) fn write_to_buffer(&mut self, df: &DataFrame) -> PolarsResult<Vec<u8>> {
        let options = self.options.clone();
        let worksheet = self.last_worksheet()?;

        Self::write_dataframe_internal(df, worksheet, 0, 0, &options)?;

        let buf = self.workbook.save_to_buffer()?;

        Ok(buf)
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

    /// Turn on/off the dataframe header in the exported Excel file.
    ///
    /// Turn on/off the dataframe header row in Excel table. It is on by
    /// default.
    ///
    /// TODO
    ///
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_has_header_on.png">
    ///
    /// If we set `has_header()` to `false` we can output the dataframe from the
    /// previous example without the header row:
    ///
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_has_header_off.png">
    ///
    pub fn set_header(&mut self, has_header: bool) -> &mut PolarsXlsxWriter {
        self.options.has_header = has_header;
        self.options.table.set_header_row(has_header);
        self
    }

    /// Set the Excel number format for time values.
    ///
    /// [Datetimes in Excel] are stored as f64 floats with a format used to
    /// display them. The default time format used by this library is
    /// `hh:mm:ss;@`. This method can be used to specify an alternative user
    /// defined format.
    ///
    /// [Datetimes in Excel]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ExcelDateTime.html#datetimes-in-excel
    ///
    /// TODO
    ///
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_time_format.png">
    ///
    pub fn set_time_format(&mut self, format: impl Into<Format>) -> &mut PolarsXlsxWriter {
        self.options.time_format = format.into();
        self
    }

    /// Set the Excel number format for date values.
    ///
    /// [Datetimes in Excel] are stored as f64 floats with a format used to
    /// display them. The default date format used by this library is
    /// `yyyy-mm-dd;@`. This method can be used to specify an alternative user
    /// defined format.
    ///
    /// [Datetimes in Excel]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ExcelDateTime.html#datetimes-in-excel
    ///
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This example
    /// demonstrates how to change the default format for Polars date types.
    ///
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/excelwriter_date_format.png">
    ///
    pub fn set_date_format(&mut self, format: impl Into<Format>) -> &mut PolarsXlsxWriter {
        self.options.date_format = format.into();
        self
    }

    /// Set the Excel number format for datetime values.
    ///
    /// [Datetimes in Excel] are stored as f64 floats with a format used to
    /// display them. The default datetime format used by this library is
    /// `yyyy-mm-dd hh:mm:ss`. This method can be used to specify an alternative
    /// user defined format.
    ///
    /// [Datetimes in Excel]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ExcelDateTime.html#datetimes-in-excel
    ///
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// example demonstrates how to change the default format for Polars
    /// datetime types.
    ///
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_datetime_format.png">
    ///
    pub fn set_datetime_format(&mut self, format: impl Into<Format>) -> &mut PolarsXlsxWriter {
        self.options.datetime_format = format.into();
        self
    }

    /// Set the Excel number format for floats.
    ///
    /// Set the Excel number format for f32/f64 float types using an Excel
    /// number format string. These format strings can be obtained for the
    /// Format Cells -> Number dialog in Excel.
    ///
    /// See all the [Number Format Categories] and subsequent sections in the
    /// `rust_xlsxwriter` documentation.
    ///
    /// [Number Format Categories]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html#number-format-categories
    ///
    /// Note, the numeric values aren't truncated in Excel, this option just
    /// controls the display of the number.
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting an Excel number format for floats.
    ///
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_float_format.png">
    ///
    pub fn set_float_format(&mut self, format: impl Into<Format>) -> &mut PolarsXlsxWriter {
        self.options.float_format = format.into();
        self
    }

    /// Set the Excel number precision for floats.
    ///
    /// Set the number precision of all floats exported from the dataframe to
    /// Excel. The precision is converted to an Excel number format (see
    /// [`set_float_format()`](PolarsXlsxWriter::set_float_format) above), so for
    /// example 3 is converted to the Excel format `0.000`.
    ///
    /// The precision should be in the Excel range 1-30.
    ///
    /// Note, the numeric values aren't truncated in Excel, this option just
    /// controls the display of the number.
    ///
    ///
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// example demonstrates how to set the precision of the float output.
    /// Setting the precision to 3 is equivalent to an Excel number format of
    /// `0.000`.
    ///
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_float_precision.png">
    ///
    pub fn set_float_precision(&mut self, precision: usize) -> &mut PolarsXlsxWriter {
        if (1..=30).contains(&precision) {
            let precision = "0".repeat(precision);
            self.options.float_format = Format::new().set_num_format(format!("0.{precision}"));
        }
        self
    }

    /// Replace Null values in the exported dataframe with string values.
    ///
    /// By default Null values in a dataframe aren't exported to Excel and will
    /// appear as empty cells. If you wish you can specify a string such as
    /// "Null", "NULL" or "N/A" as an alternative.
    ///
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting a value for Null values in the dataframe. The
    /// default is to write them as blank cells.
    ///
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_null_values.png">
    ///
    pub fn set_null_value(&mut self, null_value: impl Into<String>) -> &mut PolarsXlsxWriter {
        self.options.null_string = Some(null_value.into());
        self
    }

    /// Simulate autofit for columns in the dataframe output.
    ///
    /// Use a simulated autofit to adjust dataframe columns to the maximum
    /// string or number widths.
    ///
    /// There are several limitations to this autofit method, see the
    /// `rust_xlsxwriter` docs on [`worksheet.autofit()`] for details.
    ///
    /// [`worksheet.autofit()`]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.autofit
    ///
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// example demonstrates autofitting column widths in the output worksheet.
    ///
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_autofit.png">
    ///
    /// TODO
    pub fn set_autofit(&mut self, use_autofit: bool) -> &mut PolarsXlsxWriter {
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
