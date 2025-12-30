use std::sync::Arc;

use eyre::Context;
use tokio::sync::{mpsc, oneshot};

pub struct PsExecutor {
    pub tx: mpsc::Sender<(String, oneshot::Sender<std::result::Result<String, eyre::Report>>)>,
}

impl PsExecutor {
    pub fn new() -> Arc<Self> {
        let (tx, mut rx) = tokio::sync::mpsc::channel(8);
        let executor = Arc::new(PsExecutor { tx });
        let worker = executor.clone();

        tokio::spawn(async move {
            while let Some((command, responder)) = rx.recv().await {
                let result = tokio::process::Command::new("powershell.exe")
                    .args(["-NoProfile", "-Command", &command])
                    .output()
                    .await
                    .context("Powershell spawn failed")
                    .and_then(|out| {
                        if out.status.success() {
                            Ok(String::from_utf8_lossy(&out.stdout).trim().to_string())
                        } else {
                            Err(eyre::eyre!(
                                "Powershell command failed: {}",
                                String::from_utf8_lossy(&out.stderr)
                            ))
                        }
                    });
                let _ = responder.send(result);
            }
        });

        worker
    }

    pub async fn execute(&self, command: String) -> std::result::Result<String, eyre::Report> {
        let (responder, receiver) = oneshot::channel::<std::result::Result<String, eyre::Report>>();
        self.tx
            .send((command, responder))
            .await
            .context("Failed to send command to executor")?;
        receiver.await?
    }
}
