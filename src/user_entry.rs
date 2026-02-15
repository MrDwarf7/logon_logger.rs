use std::ops::Deref;

use calamine::Data;
use chrono::{DateTime, Local};

use crate::workstation::WorkStationEntry;
use crate::{ExcelLoggable, FieldLengsths, HasDateTime, Result};

// TODO: [trait] : Better to do this via like, S: FromStr or Into<str> or something
// TODO: [overlap] : Severe overlap between `workstationentry::WorkStationEntry` and this struct
#[derive(Clone)]
pub struct UserEntry {
    pub workstation_entry: WorkStationEntry,
    // This did originally house a bunch of fields,
    // but the data is practically identical to WorkStationEntry, so we just wrap it instead of
    // duplicate the fields.
}

impl From<WorkStationEntry> for UserEntry {
    // Caller can decide if they want to construct/clone as Ws entry
    fn from(entry: WorkStationEntry) -> Self {
        Self {
            workstation_entry: entry,
        }
    }
}

impl Deref for UserEntry {
    type Target = WorkStationEntry;

    fn deref(&self) -> &Self::Target {
        &self.workstation_entry
    }
}

impl ExcelLoggable for UserEntry {
    const COLUMNS: &'static [&'static str] = &[
        "ComputerName",
        "UserOU",
        "DateTime",
        "Period",
        "Description",
        "WS_OU",
        "OSVersion",
        "Model",
        "OS",
        "Full_OU",
        "Make",
        "UUID",
        "Serial_Number",
    ];

    fn write_entry(&self, ws: &mut rust_xlsxwriter::worksheet::Worksheet, row: u32) -> Result<()> {
        self.workstation_entry.write_entry(ws, row)
    }

    fn parse_row(row: &[Data]) -> Option<Self> {
        Some(Self {
            workstation_entry: WorkStationEntry::parse_row(row)?,
        })
    }

    fn excel_date_to_chrono(serial: f64) -> DateTime<Local> {
        WorkStationEntry::excel_date_to_chrono(serial)
    }
}

impl HasDateTime for UserEntry {
    fn date_time(&self) -> DateTime<Local> {
        self.workstation_entry.date_time()
    }
}

impl FieldLengsths for UserEntry {
    fn field_lengths(&self) -> Vec<usize> {
        self.workstation_entry.field_lengths()
    }
}
