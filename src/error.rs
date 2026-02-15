#[derive(thiserror::Error, Debug)]
pub enum Error {
    #[error("Generic error handler: {0}")]
    Generic(String),

    #[error("IO Error: {0}")]
    IoError(#[from] std::io::Error),

    #[error("Workstation error: {0}")]
    WorkStationError(#[from] rust_xlsxwriter::XlsxError),

    #[error("Tokio task join error: {0}")]
    TokioJoinError(#[from] tokio::task::JoinError),

    #[error("Calamine XLSX error: {0}")]
    CalamaineXlsxError(#[from] calamine::XlsxError),

    #[error("Environment variable error: {0}")]
    EnvironVarError(#[from] std::env::VarError),

    #[cfg(target_os = "windows")]
    #[error("Wmi error: {0}")]
    WmiError(#[from] wmi::WMIError),
}
