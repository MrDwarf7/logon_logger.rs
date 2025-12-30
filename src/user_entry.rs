use calamine::{Data, DataType};
use chrono::{DateTime, Local, TimeZone};
use rust_xlsxwriter::ExcelDateTime;

use crate::workstation::WorkStationEntry;
use crate::{Error, ExcelLoggable, FieldLengsths, HasDateTime, Result};

// TODO: @trait : Better to do this via like, S: FromStr or Into<str> or something
// TODO: @overlap : Severe overlap between `workstationentry::WorkStationEntry` and this struct
#[derive(Clone)]
pub struct UserEntry {
    pub user_ou:       String,
    pub computer_name: String,
    pub date_time:     DateTime<Local>,
    pub period:        String,
    pub description:   String,
    pub ws_ou:         String,
    pub os_version:    String,
    pub model:         String,
    pub os:            String,
    pub full_ou:       String,
    pub make:          String,
    pub uuid:          String,
    pub serial_number: String,
}

impl From<&WorkStationEntry> for UserEntry {
    fn from(entry: &WorkStationEntry) -> Self {
        Self {
            user_ou:       entry.user_ou.clone(),
            computer_name: entry.username.clone(),
            date_time:     entry.date_time,
            period:        entry.period.clone(),
            description:   entry.description.clone(),
            ws_ou:         entry.ws_ou.clone(),
            os_version:    entry.os_version.clone(),
            model:         entry.model.clone(),
            os:            entry.os.clone(),
            full_ou:       entry.full_ou.clone(),
            make:          entry.make.clone(),
            uuid:          entry.uuid.clone(),
            serial_number: entry.serial_number.clone(),
        }
    }
}

impl ExcelLoggable for UserEntry {
    const COLUMNS: &'static [&'static str] = &[
        "UserOU",
        "ComputerName",
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
        let edt = ExcelDateTime::parse_from_str(&self.date_time.to_rfc3339())
            .map_err(|e| Error::Generic(format!("Failed to parse date time: {}", e)))?;

        ws.write_string(row, 0, &self.user_ou)?;
        ws.write_string(row, 1, &self.computer_name)?;
        ws.write_datetime(row, 2, edt)?;
        ws.write_string(row, 3, &self.period)?;
        ws.write_string(row, 4, &self.description)?;
        ws.write_string(row, 5, &self.ws_ou)?;
        ws.write_string(row, 6, &self.os_version)?;
        ws.write_string(row, 7, &self.model)?;
        ws.write_string(row, 8, &self.os)?;
        ws.write_string(row, 9, &self.full_ou)?;
        ws.write_string(row, 10, &self.make)?;
        ws.write_string(row, 11, &self.uuid)?;
        ws.write_string(row, 12, &self.serial_number)?;
        Ok(())
    }

    fn parse_row(row: &[Data]) -> Option<Self> {
        if row.len() < 13 {
            return None;
        }

        let dt = if row[2].get_float().is_some() {
            Self::excel_date_to_chrono(row[2].get_float()?)
        } else {
            return None;
        };

        Some(UserEntry {
            user_ou:       row[0].as_string()?,
            computer_name: row[1].as_string()?,
            date_time:     dt,
            period:        row[3].as_string()?,
            description:   row[4].as_string()?,
            ws_ou:         row[5].as_string()?,
            os_version:    row[6].as_string()?,
            model:         row[7].as_string()?,
            os:            row[8].as_string()?,
            full_ou:       row[9].as_string()?,
            make:          row[10].as_string()?,
            uuid:          row[11].as_string()?,
            serial_number: row[12].as_string()?,
        })
    }

    fn excel_date_to_chrono(serial: f64) -> DateTime<Local> {
        let unix = ((serial - 25569.0) * 86400.0).round() as i64;
        chrono::Local
            .timestamp_opt(unix, 0)
            .single()
            .expect("Invalid timestamp")
            .with_timezone(&chrono::Local)
    }
}

impl HasDateTime for UserEntry {
    fn date_time(&self) -> DateTime<Local> {
        self.date_time
    }
}

impl FieldLengsths for UserEntry {
    fn field_lengths(&self) -> Vec<usize> {
        vec![
            self.user_ou.len(),
            self.computer_name.len(),
            self.period.len(),
            self.description.len(),
            self.ws_ou.len(),
            self.os_version.len(),
            self.model.len(),
            self.os.len(),
            self.full_ou.len(),
            self.make.len(),
            self.uuid.len(),
            self.serial_number.len(),
        ]
    }
}
