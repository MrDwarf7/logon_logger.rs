use calamine::{Data, DataType};
use chrono::{DateTime, Local, TimeZone};
use rust_xlsxwriter::ExcelDateTime;
use rust_xlsxwriter::worksheet::Worksheet;

use crate::collect::{BaseInfo, HardwareInfo, OsInfo};
use crate::period::{PERIODS, get_current_period};
use crate::{Error, ExcelLoggable, FieldLengsths, HasDateTime, Result};

// TODO: [trait] : Better to do this via like, S: FromStr or Into<str> or something
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

impl From<(BaseInfo, HardwareInfo, OsInfo, DateTime<Local>)> for WorkStationEntry {
    fn from(value: (BaseInfo, HardwareInfo, OsInfo, DateTime<Local>)) -> Self {
        let (base, hardware, os, now) = value;
        // let now = chrono::Local::now();
        let period_value = get_current_period(&now, &PERIODS).unwrap_or_else(|_| "Unknown".to_string());
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
        #[allow(unused_imports)]
        use rust_xlsxwriter::Format;

        let fields = [
            &self.username,                   // 0
            &self.user_ou,                    // 1
            &String::from("__placeholder__"), // 2 &edt placeholder, overwritten by DT formatted col
            &self.period,                     // 3
            &self.description,                // 4
            &self.ws_ou,                      // 5
            &self.os_version,                 // 6
            &self.model,                      // 7
            &self.os,                         // 8
            &self.full_ou,                    // 9
            &self.make,                       // 10
            &self.uuid,                       // 11
            &self.serial_number,              // 12
        ];

        fields.iter().enumerate().for_each(|(i, field)| {
            ws.write_string(row, i as u16, *field).unwrap_or_else(|e| {
                panic!("Failed to write field {}: {}", Self::COLUMNS[i], e);
            });
        });

        // convert the `edt` datetime from the string insertion format,
        // back to a datetime by.... writing it again as an actual datetime

        let edt = ExcelDateTime::parse_from_str(&self.date_time.to_rfc3339())
            .map_err(|e| Error::Generic(format!("Failed to parse date time: {}", e)))?;

        // HACK: [dirty] : Extremely dirty way to handle this lol...
        // Can be cleaned up if we're okay with having the DT at the end of the columns in the sheet
        ws.write_datetime(row, 2, &edt)?;

        Ok(())
    }

    fn parse_row(row: &[Data]) -> Option<Self> {
        // NOTE: ?[maybe] : rows.len ??? or is it via columns len?
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
