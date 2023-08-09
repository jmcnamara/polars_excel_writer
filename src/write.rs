// excel_writer - A Polars extension to serialize dataframes to Excel xlsx files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023, John McNamara, jmcnamara@cpan.org

use std::io::Write;

use polars::prelude::*;

use crate::PolarsXlsxWriter;

pub struct ExcelWriter<W: Write> {
    buffer: W,
    //header: bool,
}

impl<W> SerWriter<W> for ExcelWriter<W>
where
    W: Write,
{
    fn new(buffer: W) -> Self {
        ExcelWriter {
            buffer,
            //header: true,
        }
    }

    fn finish(&mut self, df: &mut DataFrame) -> PolarsResult<()> {
        let xlsx_writer = PolarsXlsxWriter::new();
        let bytes = xlsx_writer.write_buffer(df).unwrap();
        self.buffer.write_all(&bytes).unwrap();

        Ok(())
    }
}
