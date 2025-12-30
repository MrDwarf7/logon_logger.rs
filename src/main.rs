mod collect;
mod error;
mod executor;
mod period;
mod prelude;

mod user_entry;
mod workstation;

use std::cmp::Reverse;
use std::fmt::Display;
use std::path::PathBuf;
use std::str::FromStr;

use calamine::{Data, DataType, Reader, Xlsx};
use chrono::{DateTime, Local};
use rust_xlsxwriter::workbook::Workbook;
use rust_xlsxwriter::worksheet::Worksheet;
use rust_xlsxwriter::{Format, Table, TableStyle};

pub use self::prelude::{Error, Result, W};
use crate::collect::{collect_base_info, collect_hardware, collect_os_info};
use crate::executor::PsExecutor;
use crate::period::{PERIODS, get_current_period};
use crate::user_entry::UserEntry;
use crate::workstation::WorkStationEntry;

pub trait ExcelLoggable: Sized + Clone {
    const COLUMNS: &'static [&'static str];

    fn write_entry(&self, ws: &mut Worksheet, row: u32) -> Result<()>; // as rust_xlsxwriter::Resultlsxwriter

    fn parse_row(row: &[Data]) -> Option<Self>;
    // where
    // D: AsRef<dyn DataType>;

    fn excel_date_to_chrono(serial: f64) -> DateTime<Local>;
}

pub trait HasDateTime {
    fn date_time(&self) -> DateTime<Local>;
}

pub trait FieldLengsths {
    fn field_lengths(&self) -> Vec<usize>;
}

const WORKSHEET_NAME: &str = "Logons";

async fn append_log<E, S>(base_path: S, file_base: S, new_entry: E) -> Result<()>
where
    E: ExcelLoggable + HasDateTime + Send + 'static,
    E: FieldLengsths,
    S: AsRef<str> + Display,
{
    let base_path = base_path.as_ref();
    let file_base = file_base.as_ref();
    let path = PathBuf::from(base_path).join(format!("{file_base}.xlsx"));

    let new_path = path.clone();

    let entries: Vec<E> = tokio::task::spawn_blocking(move || -> Result<Vec<E>> {
        let mut existing = vec![];
        if path.exists() {
            let mut wb: Xlsx<_> = calamine::open_workbook(&path)?;
            if let Ok(range) = wb.worksheet_range(WORKSHEET_NAME) {
                for r in range.rows().skip(1) {
                    if let Some(e) = E::parse_row(r) {
                        existing.push(e);
                    }
                }
            }
        }
        existing.push(new_entry);
        existing.sort_by_key(|e| Reverse(e.date_time()));
        Ok(existing)
    })
    .await??;

    let path = new_path.clone();

    tokio::task::spawn_blocking(move || -> Result<()> {
        std::fs::create_dir(path.parent().unwrap())?;
        let mut workbook = Workbook::new();
        let ws = workbook.add_worksheet();
        ws.set_name(WORKSHEET_NAME)?;

        let bold = Format::new().set_bold();
        let date_fmt = Format::new().set_num_format("yyyy/mm/dd hh:mm AM/PM");
        ws.set_column_format(2, &date_fmt)?;

        for (c, h) in E::COLUMNS.iter().enumerate() {
            ws.write_string_with_format(0, c as u16, *h, &bold)?;
        }

        let mut widths: Vec<usize> = E::COLUMNS.iter().map(|h| h.len()).collect();
        for e in &entries {
            let values = e.field_lengths();
            for (i, v) in values.iter().enumerate() {
                if *v > widths[i] {
                    widths[i] = *v;
                }
            }
            // first 2 are handled in impl. spec if needed (?)
        }

        for (i, w) in widths.iter().enumerate() {
            ws.set_column_width(i as u16, (*w + 2) as f64)?; // +2 for padding
        }

        // Data rows
        for (i, e) in entries.iter().enumerate() {
            e.write_entry(ws, (i + 1) as u32)?;
        }

        // table + formatting
        if !entries.is_empty() {
            let table = Table::new().set_style(TableStyle::Medium9);
            ws.add_table(0, 0, entries.len() as u32, E::COLUMNS.len() as u16 - 1, &table)?;
            ws.autofilter(0, 0, entries.len() as u32, E::COLUMNS.len() as u16 - 1)?;
        }
        ws.set_freeze_panes(1, 0)?;

        workbook.save(path)?;
        Ok(())
    })
    .await??;

    Ok(())
}

#[tokio::main]
async fn main() -> Result<()> {
    let now = Local::now();
    let executor = PsExecutor::new();

    let period = get_current_period(&now, &PERIODS);

    let (base_info, hardware_info, os_info) =
        tokio::try_join!(collect_base_info(&executor), collect_hardware(), collect_os_info())?;

    // let base_info = collect_base_info(&executor).await?;
    // let hardware_info = collect_hardware().await?;
    // let os_info = collect_os_info().await?;

    let ws = WorkStationEntry::from((base_info, hardware_info, os_info));
    let user_entry = UserEntry::from(&ws);

    const WS_BASE_PATH: &str = r"\\Server\LogonLogger$\Logs\ComputerNEW";
    const USER_BASE_PATH: &str = r"\\Server\LogonLogger$\Logs\UserNEW";

    let today = chrono::NaiveDate::from_str(&now.format("%Y-%m-%d").to_string())
        .map_err(|e| Error::Generic(format!("Failed to format date: {}", e)))?;

    let workstation_log = String::from(format!("workstation_log_{}", today));
    let user_log = String::from(format!("user_log_{}", today));

    tokio::try_join!(
        append_log(WS_BASE_PATH, &workstation_log, ws),
        append_log(USER_BASE_PATH, &user_log, user_entry)
    )?;

    Ok(())
}
