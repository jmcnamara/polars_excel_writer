use std::fs::File;

use chrono::prelude::*;
use polars::prelude::*;
use polars_excel_writer::ExcelWriter;

fn main() {
    let mut df: DataFrame = df!(
        "String" => &["North", "South", "East", "West", "All"],
        "Integer" => &[1, 2, 3, 4, 5],
        "Datetime" => &[
            NaiveDate::from_ymd_opt(2022, 1, 1).unwrap().and_hms_opt(1, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 2).unwrap().and_hms_opt(2, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 3).unwrap().and_hms_opt(3, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 4).unwrap().and_hms_opt(4, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 5).unwrap().and_hms_opt(5, 0, 0).unwrap(),
        ],
        "Date" => &[
            NaiveDate::from_ymd_opt(2022, 1, 1).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 2).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 3).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 4).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 5).unwrap(),
        ],
        "Time" => &[
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
        ],
        "Float" => &[4.0, 5.0, 6.0, 7.0, 8.0],
    )
    .expect("should not fail");

    let mut file = File::create("dataframe.xlsx").expect("could not create file");

    ExcelWriter::new(&mut file)
        .has_header(false)
        .with_float_precision(3)
        .with_time_format("hh::mm".to_string())
        .with_date_format("yyyy mmm dd".to_string())
        .with_datetime_format("yyyy mmm dd hh::mm".to_string())
        .finish(&mut df)
        .unwrap();
}
