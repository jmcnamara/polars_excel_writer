# polars_excel_writer

The `polars_excel_writer` crate is a library for serializing Polars dataframes
to Excel XLSX files.

It provides two interfaces for writing a dataframe to an Excel XLSX file:

- [`ExcelWriter`] a simple Excel serializer that implements the Polars
  [`SerWriter`] trait to write a dataframe to an Excel XLSX file.

- [`PolarsXlsxWriter`] a more configurable Excel serializer that more closely
  resembles the interface options provided by the Polars [`write_excel()`]
  dataframe method.

`ExcelWriter` uses `PolarsXlsxWriter` to do the Excel serialization which in
turn uses the [`rust_xlsxwriter`] crate.

[`ExcelWriter`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/struct.ExcelWriter.html
[`PolarsXlsxWriter`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/struct.PolarsXlsxWriter.html

[`SerWriter`]:
    https://docs.rs/polars/latest/polars/prelude/trait.SerWriter.html

[`CsvWriter`]:
    https://docs.rs/polars/latest/polars/prelude/struct.CsvWriter.html

[`rust_xlsxwriter`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/

[`write_excel()`]:
   https://pola-rs.github.io/polars/py-polars/html/reference/api/polars.DataFrame.write_excel.html#polars.DataFrame.write_excel

## Examples

An example of writing a Polar Rust dataframe to an Excel file using the
`ExcelWriter` interface.

```rust
use polars::prelude::*;
use chrono::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Integer" => &[1, 2, 3, 4],
        "Float" => &[4.0, 5.0, 6.0, 7.0],
        "Time" => &[
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            ],
        "Date" => &[
            NaiveDate::from_ymd_opt(2022, 1, 1).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 2).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 3).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 4).unwrap(),
            ],
        "Datetime" => &[
            NaiveDate::from_ymd_opt(2022, 1, 1).unwrap().and_hms_opt(1, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 2).unwrap().and_hms_opt(2, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 3).unwrap().and_hms_opt(3, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 4).unwrap().and_hms_opt(4, 0, 0).unwrap(),
        ],
    )
    .unwrap();

    example(&mut df).unwrap();
}

use polars_excel_writer::ExcelWriter;

fn example(mut df: &mut DataFrame) -> PolarsResult<()> {
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    ExcelWriter::new(&mut file)
        .finish(&mut df)
}
```

Output file:

<img src="https://rustxlsxwriter.github.io/images/excelwriter_intro.png">

