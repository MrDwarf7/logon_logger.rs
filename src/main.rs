mod append;
mod collect;
mod error;
mod executor;
mod period;
mod prelude;
mod user_entry;
mod workstation;
use std::str::FromStr;

use calamine::Data;
use chrono::{DateTime, Local};
use rust_xlsxwriter::worksheet::Worksheet;

use crate::append::append_log;
use crate::collect::{collect_base_info, collect_hardware, collect_os_info};
use crate::executor::PsExecutor;
pub use crate::prelude::{Error, Result, W};
use crate::user_entry::UserEntry;
use crate::workstation::WorkStationEntry;

pub trait ExcelLoggable: Sized + Clone {
    const COLUMNS: &'static [&'static str];

    fn write_entry(&self, ws: &mut Worksheet, row: u32) -> Result<()>; // as rust_xlsxwriter::Resultlsxwriter

    fn parse_row(row: &[Data]) -> Option<Self>;

    fn excel_date_to_chrono(serial: f64) -> DateTime<Local>;
}

pub trait HasDateTime {
    fn date_time(&self) -> DateTime<Local>;
}

pub trait FieldLengsths {
    fn field_lengths(&self) -> Vec<usize>;
}

// TODO: [customizability] : Move these to be changable via env vars with compile time (literal)
// fallbacks. Additionally, might consider having a config we read from at program start (on a
// network drive etc.). These would then need to be `static` instead.

const WORKSHEET_NAME: &str = "Logons";
const WS_BASE_PATH: &str = r"\\Server\LogonLogger$\Logs\ComputerNEW";
const USER_BASE_PATH: &str = r"\\Server\LogonLogger$\Logs\UserNEW";

#[tokio::main]
async fn main() -> Result<()> {
    let now = Local::now();
    let executor = PsExecutor::new();

    // let base_info = collect_base_info(&executor).await?;
    // let hardware_info = collect_hardware().await?;
    // let os_info = collect_os_info().await?;

    let today = chrono::NaiveDate::from_str(&now.format("%Y-%m-%d").to_string())
        .map_err(|e| Error::Generic(format!("Failed to format date: {}", e)))?;

    let workstation_log = format!("workstation_log_{today}");
    let user_log = format!("user_log_{today}");

    let (base_info, hardware_info, os_info) =
        tokio::try_join!(collect_base_info(&executor), collect_hardware(), collect_os_info())?;

    let ws = WorkStationEntry::from((base_info, hardware_info, os_info, now));
    let user_entry = UserEntry::from(ws.clone());

    tokio::try_join!(
        append_log(WS_BASE_PATH, &workstation_log, ws),
        append_log(USER_BASE_PATH, &user_log, user_entry)
    )?;

    Ok(())
}
