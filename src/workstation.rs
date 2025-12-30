use calamine::{Data, DataType};
use chrono::{DateTime, Local, TimeZone};
use rust_xlsxwriter::ExcelDateTime;
use rust_xlsxwriter::worksheet::Worksheet;

use crate::collect::{BaseInfo, HardwareInfo, OsInfo};
use crate::period::{PERIODS, get_current_period};
use crate::{Error, ExcelLoggable, FieldLengsths, HasDateTime, Result};

// TODO: @trait : Better to do this via like, S: FromStr or Into<str> or something
#[derive(Clone)]
pub struct WorkStationEntry {
    pub username:      String,
    pub user_ou:       String,
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

impl From<(BaseInfo, HardwareInfo, OsInfo)> for WorkStationEntry {
    fn from(value: (BaseInfo, HardwareInfo, OsInfo)) -> Self {
        let (base, hardware, os) = value;
        let now = chrono::Local::now();
        let period_value = get_current_period(&now, &PERIODS);
        Self {
            username:      base.username,
            user_ou:       base.user_ou,
            date_time:     now,
            period:        period_value,
            description:   hardware.os_description,
            ws_ou:         base.ws_ou,
            os_version:    os.os_version,
            model:         hardware.model,
            os:            os.os_name,
            full_ou:       base.full_ou,
            make:          hardware.make,
            uuid:          hardware.uuid,
            serial_number: hardware.serial_number,
        }
    }
}

impl ExcelLoggable for WorkStationEntry {
    const COLUMNS: &'static [&'static str] = &[
        "Username",
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

    fn write_entry(&self, ws: &mut Worksheet, row: u32) -> Result<()> {
        use rust_xlsxwriter::Format;

        let edt = ExcelDateTime::parse_from_str(&self.date_time.to_rfc3339())
            .map_err(|e| Error::Generic(format!("Failed to parse date time: {}", e)))?;

        ws.write_string(row, 0, &self.username)?;
        ws.write_string(row, 1, &self.user_ou)?;
        ws.write_datetime(row, 2, &edt)?;
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

        // check if the row[2] can be a float, else return None;

        let dt = if row[2].get_float().is_some() {
            Self::excel_date_to_chrono(row[2].get_float().unwrap())
        } else {
            return None;
        };

        Some(WorkStationEntry {
            username:      row[0].as_string()?,
            user_ou:       row[1].as_string()?,
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
        // let mut utc = chrono::Utc::now();
        // let mut local = chrono::Local::now();

        chrono::Local
            .timestamp_opt(unix, 0)
            .single()
            .expect("Invalid excel date")
            .with_timezone(&chrono::Local)
    }
}

impl HasDateTime for WorkStationEntry {
    fn date_time(&self) -> DateTime<Local> {
        self.date_time
    }
}

impl FieldLengsths for WorkStationEntry {
    fn field_lengths(&self) -> Vec<usize> {
        vec![
            self.username.len(),
            self.user_ou.len(),
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
