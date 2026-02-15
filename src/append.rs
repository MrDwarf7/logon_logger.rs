use std::cmp::Reverse;
use std::fmt::Display;
use std::path::PathBuf;

use calamine::{Reader as _, Xlsx};
use rust_xlsxwriter::workbook::Workbook;
use rust_xlsxwriter::{Format, Table, TableStyle};

use crate::prelude::Result;
use crate::{ExcelLoggable, FieldLengsths, HasDateTime, WORKSHEET_NAME};

pub async fn append_log<S, E>(base_path: S, file_base: S, new_entry: E) -> Result<()>
where
    S: AsRef<str> + Display,
    E: ExcelLoggable + HasDateTime + Send + 'static,
    E: FieldLengsths,
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
                // skip header
                for r in range.rows().skip(1) {
                    if let Some(e) = E::parse_row(r) {
                        // E::parse_row(r) {
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
