// Entry point for `polars_excel_writer` library.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023, John McNamara, jmcnamara@cpan.org

pub mod write;
pub mod xlsx_writer;

pub use write::*;
pub use xlsx_writer::*;
