use chrono::{DateTime, Local};

use crate::Result;
use crate::executor::PsExecutor;

#[derive(Debug, Clone)]
pub struct BaseInfo {
    pub computer_name: String,
    pub username:      String,
    pub now:           DateTime<Local>,
    pub user_ou:       String,
    pub full_ou:       String,
    pub ws_ou:         String,
}

impl BaseInfo {
    pub fn new(
        computer_name: String,
        username: String,
        now: DateTime<Local>,
        user_ou: String,
        full_ou: String,
        ws_ou: String,
    ) -> Self {
        Self {
            computer_name,
            username,
            now,
            user_ou,
            full_ou,
            ws_ou,
        }
    }
}

pub async fn collect_base_info(executor: &PsExecutor) -> Result<BaseInfo> {
    // TODO: [correctness] : We should be uppercasing these,
    // left them as is due to original script. Consider uppercasing as a way to normalise the data.
    let computer_name = std::env::var("COMPUTERNAME")
        .map_err(|_| crate::Error::Generic("COMPUTERNAME env var not found".to_string()))?; // .to_uppercase();

    let logon_name = std::env::var("USERNAME")
        .map_err(|_| crate::Error::Generic("USERNAME env var not found".to_string()))?; // .to_uppercase();

    let now = Local::now();

    let user_dn_cmd =
        format!("(Get-ADUser -Identity '{}' -Properties DistinguishedName).DistinguishedName", logon_name);
    let comp_dn_cmd = format!(
        "(Get-ADComputer -Identity '{}' -Properties DistinguishedName).DistinguishedName",
        computer_name
    );

    let (user_dn, comp_dn) = tokio::try_join!(executor.execute(user_dn_cmd), executor.execute(comp_dn_cmd))
        .map_err(|e| crate::Error::Generic(format!("Failed to get DN: {}", e)))?;

    // HACK: [brittle] : This is _extremely_ brittle,
    // if the structure of the output changes, this _WILL_ break.
    // perhaps Regex? (though questionable...), parsing it properly
    // would be much safer and resilient...

    let username = user_dn
        .split(',')
        .next()
        .and_then(|cn| cn.split('=').nth(1))
        .unwrap_or(&logon_name)
        .to_string();

    let user_ou = user_dn
        .split(',')
        .nth(1)
        .and_then(|ou| ou.split('_').nth(1))
        .unwrap_or("Unknown")
        .to_string();

    let full_ou = comp_dn.split(',').nth(1).unwrap_or("Unknown").to_string();

    // let ws_ou = full_ou.split('_').last().unwrap_or("Unknown").to_string();
    let ws_ou = full_ou.split('_').next_back().unwrap_or("Unknown").to_string();

    Ok(BaseInfo::new(computer_name, username, now, user_ou, full_ou, ws_ou))
}

#[derive(Debug, Clone)]
pub struct HardwareInfo {
    pub make:           String,
    pub model:          String,
    pub uuid:           String,
    pub serial_number:  String,
    pub os_description: String,
}

impl HardwareInfo {
    pub fn new(
        make: String,
        model: String,
        uuid: String,
        serial_number: String,
        os_description: String,
    ) -> Self {
        Self {
            make,
            model,
            uuid,
            serial_number,
            os_description,
        }
    }
}

pub async fn collect_hardware() -> Result<HardwareInfo> {
    #[cfg(target_os = "windows")]
    return tokio::task::spawn_blocking(|| {
        use serde::{Deserialize, Serialize};
        use winreg::RegKey;
        use winreg::enums::*;
        use wmi::{COMLibrary, WMIConnection};

        // let com = COMLibrary::new()?;
        let wmi_con = WMIConnection::new()?;

        #[derive(Debug, Deserialize)]
        struct CS {
            Manufacturer: Option<String>,
            Model:        Option<String>,
        }
        let cs: Vec<CS> = wmi_con.raw_query("SELECT Manufacturer, Model FROM Win32_ComputerSystem")?;
        let make = cs
            .first()
            .and_then(|c| c.Manufacturer.clone())
            .unwrap_or_default();
        let model = cs.first().and_then(|c| c.Model.clone()).unwrap_or_default();

        #[derive(Debug, Deserialize)]
        struct CSP {
            UUID: Option<String>,
        }
        let uuid: Vec<CSP> = wmi_con.raw_query("SELECT UUID FROM Win32_ComputerSystemProduct")?;
        let uuid = uuid.first().and_then(|c| c.UUID.clone()).unwrap_or_default();

        #[derive(Debug, Deserialize)]
        struct Bios {
            SerialNumber: Option<String>,
        }
        let serial: Vec<Bios> = wmi_con.raw_query("SELECT SerialNumber FROM Win32_BIOS")?;
        let serial_number = serial
            .first()
            .and_then(|c| c.SerialNumber.clone())
            .unwrap_or_default();

        #[derive(Debug, Deserialize)]
        struct OSDesc {
            Description: Option<String>,
        }
        let desc: Vec<OSDesc> = wmi_con.raw_query("SELECT Description FROM Win32_OperatingSystem")?;
        let description = desc
            .first()
            .and_then(|c| c.Description.clone())
            .unwrap_or_default();

        Ok(HardwareInfo::new(make, model, uuid, serial_number, description))
    })
    .await?;

    #[cfg(not(target_os = "windows"))]
    {
        unimplemented!("Hardware collection is only implemented for Windows");
    }
}

#[derive(Debug, Clone)]
pub struct OsInfo {
    pub os_version: String,
    pub os_name:    String,
}

impl OsInfo {
    pub fn new(os_version: String, os_name: String) -> Self {
        Self { os_version, os_name }
    }
}

pub async fn collect_os_info() -> Result<OsInfo> {
    #[cfg(target_os = "windows")]
    return tokio::task::spawn_blocking(|| {
        let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);
        let cur = hklm.open_sub_key_with_flags(r"SOFTWARE\Microsoft\Windows NT\CurrentVersion", KEY_READ)?;
        let os: String = cur.get_value("ProductName")?;
        let os_version: String = cur.get_value("DisplayVersion").unwrap_or_default();
        Ok(OsInfo::new(os_version, os))
    })
    .await?;

    #[cfg(not(target_os = "windows"))]
    {
        unimplemented!("OS info collection is only implemented for Windows");
    }
}
