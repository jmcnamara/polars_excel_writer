// excel_writer - A Polars extension to serialize dataframes to Excel xlsx files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023, John McNamara, jmcnamara@cpan.org

use std::io::Write;

use polars::prelude::*;

use crate::PolarsXlsxWriter;

pub struct ExcelWriter<W>
where
    W: Write,
{
    writer: W,
    pub(crate) has_header: bool,
    pub(crate) date_format: String,
    pub(crate) time_format: String,
    pub(crate) null_string: String,
    pub(crate) float_format: String,
    pub(crate) datetime_format: String,
}

impl<W> SerWriter<W> for ExcelWriter<W>
where
    W: Write,
{
    fn new(buffer: W) -> Self {
        ExcelWriter {
            writer: buffer,
            has_header: true,
            date_format: String::new(),
            null_string: String::new(),
            time_format: String::new(),
            float_format: String::new(),
            datetime_format: String::new(),
        }
    }

    fn finish(&mut self, df: &mut DataFrame) -> PolarsResult<()> {
        let xlsx_writer = PolarsXlsxWriter::new_from_excel_writer(self);
        let bytes = xlsx_writer.write_buffer(df).unwrap();
        self.writer.write_all(&bytes).unwrap();

        Ok(())
    }
}

impl<W> ExcelWriter<W>
where
    W: Write,
{
    /// TODO
    pub fn has_header(mut self, has_header: bool) -> Self {
        self.has_header = has_header;
        self
    }

    /// TODO
    pub fn with_date_format(mut self, format: String) -> Self {
        self.date_format = format;
        self
    }

    /// TODO
    pub fn with_time_format(mut self, format: String) -> Self {
        self.time_format = format;
        self
    }

    /// TODO
    pub fn with_datetime_format(mut self, format: String) -> Self {
        self.datetime_format = format;
        self
    }

    /// TODO
    pub fn with_float_format(mut self, format: String) -> Self {
        self.float_format = format;
        self
    }

    /// TODO
    pub fn with_float_precision(mut self, precision: usize) -> Self {
        if precision > 0 {
            let precision = "0".repeat(precision);
            self.float_format = format!("0.{precision}");
        }
        self
    }

    /// TODO
    pub fn with_null_value(mut self, null_value: String) -> Self {
        self.null_string = null_value;
        self
    }
}
