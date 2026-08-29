use tauri::{
    menu::{Menu, MenuItem, PredefinedMenuItem},
    tray::{MouseButton, MouseButtonState, TrayIconBuilder, TrayIconEvent},
    Emitter, Manager,
};
use tauri_plugin_global_shortcut::{GlobalShortcutExt, ShortcutState};
use calamine::{open_workbook_auto, Data, DataType, Reader};
use chrono::{Datelike, Duration, NaiveDate, TimeZone, Utc};
use regex::Regex;
use base64::{engine::general_purpose::STANDARD as BASE64, Engine as _};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::sync::{Arc, RwLock};

#[derive(Clone, Default)]
struct OutlookBridgeState(Arc<RwLock<Vec<String>>>);

#[derive(serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct OutlookCaptureQuery {
    key: String,
    item_id: Option<String>,
    subject: Option<String>,
    sender_name: Option<String>,
    sender_email: Option<String>,
    recipients: Option<String>,
    received_at: Option<String>,
    excerpt: Option<String>,
    attachments: Option<String>,
}

#[derive(serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct OutlookAttachmentPayload {
    name: String,
    content: String,
}

#[derive(Clone, serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct CapturedAttachment {
    id: String,
    name: String,
    path: String,
    size: usize,
}

#[derive(Clone, serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct CapturedEmail {
    id: String,
    outlook_item_id: String,
    subject: String,
    sender_name: String,
    sender_email: String,
    recipients: String,
    received_at: String,
    excerpt: String,
    attachments: Vec<CapturedAttachment>,
    captured_at: String,
    processed: bool,
}

fn safe_file_name(value: &str) -> String {
    let name = Path::new(value).file_name().and_then(|name| name.to_str()).unwrap_or("Вложение");
    let cleaned = name.chars().map(|character| match character {
        '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*' => '_',
        _ => character,
    }).collect::<String>();
    let cleaned = cleaned.trim().trim_end_matches(&['.', ' '][..]).to_string();
    if cleaned.is_empty() { "Вложение".to_string() } else { cleaned }
}

fn available_destination(directory: &Path, file_name: &str) -> PathBuf {
    let initial = directory.join(file_name);
    if !initial.exists() { return initial; }
    let path = Path::new(file_name);
    let stem = path.file_stem().and_then(|value| value.to_str()).unwrap_or("Вложение");
    let extension = path.extension().and_then(|value| value.to_str());
    for index in 2..1000 {
        let candidate = match extension {
            Some(extension) => directory.join(format!("{stem} ({index}).{extension}")),
            None => directory.join(format!("{stem} ({index})")),
        };
        if !candidate.exists() { return candidate; }
    }
    directory.join(format!("{}-{}", Utc::now().timestamp_millis(), file_name))
}

fn store_outlook_attachments(app: &tauri::AppHandle, capture_id: &str, encoded: Option<&str>) -> Result<Vec<CapturedAttachment>, String> {
    let payloads = match encoded.filter(|value| !value.trim().is_empty()) {
        Some(value) => serde_json::from_str::<Vec<OutlookAttachmentPayload>>(value).map_err(|_| "Не удалось прочитать список вложений Outlook".to_string())?,
        None => return Ok(Vec::new()),
    };
    if payloads.is_empty() { return Ok(Vec::new()); }
    let directory = app.path().app_data_dir().map_err(|error| format!("Не удалось определить папку данных: {error}"))?.join("outlook-attachments").join(capture_id);
    std::fs::create_dir_all(&directory).map_err(|error| format!("Не удалось создать папку вложений: {error}"))?;
    let mut stored = Vec::new();
    for (index, payload) in payloads.into_iter().enumerate() {
        let bytes = BASE64.decode(payload.content.trim()).map_err(|_| format!("Вложение «{}» передано в неверном формате", payload.name))?;
        if bytes.len() > 50 * 1024 * 1024 { return Err(format!("Вложение «{}» превышает 50 МБ", payload.name)); }
        let name = safe_file_name(&payload.name);
        let path = available_destination(&directory, &name);
        std::fs::write(&path, &bytes).map_err(|error| format!("Не удалось сохранить вложение «{name}»: {error}"))?;
        stored.push(CapturedAttachment { id: format!("attachment-{index}"), name: path.file_name().unwrap_or_default().to_string_lossy().into_owned(), path: path.to_string_lossy().into_owned(), size: bytes.len() });
    }
    Ok(stored)
}

#[tauri::command]
fn set_outlook_bridge_key(key: String, state: tauri::State<'_, OutlookBridgeState>) -> Result<(), String> {
    let key = key.trim().to_string();
    if key.len() < 16 { return Err("Код сопряжения слишком короткий".to_string()); }
    let mut keys = state.0.write().map_err(|_| "Не удалось обновить код сопряжения")?;
    keys.retain(|stored| stored != &key);
    keys.push(key);
    while keys.len() > 4 { keys.remove(0); }
    Ok(())
}

fn start_outlook_bridge(app: tauri::AppHandle, state: OutlookBridgeState) {
    std::thread::spawn(move || {
        let server = match tiny_http::Server::http("127.0.0.1:17832") {
            Ok(server) => server,
            Err(error) => { eprintln!("failed to start Outlook bridge: {error}"); return; }
        };
        for mut request in server.incoming_requests() {
            let url = request.url().to_string();
            let (path, query) = url.split_once('?').unwrap_or((&url, ""));
            if path != "/capture" {
                let _ = request.respond(tiny_http::Response::from_string("Not found").with_status_code(404));
                continue;
            }
            let mut body = String::new();
            let encoded = if request.method() == &tiny_http::Method::Post {
                if request.as_reader().read_to_string(&mut body).is_err() {
                    let _ = request.respond(tiny_http::Response::from_string("Не удалось прочитать данные письма.").with_status_code(400));
                    continue;
                }
                body.as_str()
            } else {
                query
            };
            let parsed = serde_urlencoded::from_str::<OutlookCaptureQuery>(encoded);
            let authorized = parsed.as_ref().ok()
                .and_then(|capture| state.0.read().ok().map(|keys| keys.iter().any(|key| key == capture.key.trim())))
                .unwrap_or(false);
            let response = match parsed {
                Ok(capture) if authorized => {
                    let id = format!("outlook-{}", Utc::now().timestamp_millis());
                    let attachments = store_outlook_attachments(&app, &id, capture.attachments.as_deref());
                    match attachments {
                    Ok(attachments) => {
                    let attachment_count = attachments.len();
                    let email = CapturedEmail {
                        id,
                        outlook_item_id: capture.item_id.unwrap_or_default(),
                        subject: capture.subject.unwrap_or_default(),
                        sender_name: capture.sender_name.unwrap_or_default(),
                        sender_email: capture.sender_email.unwrap_or_default(),
                        recipients: capture.recipients.unwrap_or_default(),
                        received_at: capture.received_at.unwrap_or_default(),
                        excerpt: capture.excerpt.unwrap_or_default(),
                        attachments,
                        captured_at: Utc::now().to_rfc3339(),
                        processed: false,
                    };
                    let _ = app.emit("outlook-email-captured", email);
                    show_window(&app, "main");
                    tiny_http::Response::from_string(format!("<html><meta charset=\"utf-8\"><body style=\"font:16px Segoe UI;padding:32px\"><h2>Письмо добавлено в MyPlanner</h2><p>Получено вложений: {attachment_count}. Это окно можно закрыть.</p><script>setTimeout(()=>window.close(),1800)</script></body></html>")).with_status_code(200)
                    }
                    Err(message) => tiny_http::Response::from_string(message).with_status_code(400),
                    }
                }
                Ok(_) => tiny_http::Response::from_string("Неверный код сопряжения. Откройте настройки Outlook в MyPlanner.").with_status_code(403),
                Err(_) => tiny_http::Response::from_string("Некорректные данные письма.").with_status_code(400),
            };
            let content_type = tiny_http::Header::from_bytes(&b"Content-Type"[..], &b"text/html; charset=utf-8"[..]).expect("valid content type header");
            let _ = request.respond(response.with_header(content_type));
        }
    });
}

const YOUTRACK_BASE_URL: &str = "https://youtrack.advalange.com";

#[derive(Clone, serde::Serialize, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct WorklogPreviewEntry {
    issue_key: String,
    minutes: i64,
    work_date: String,
}

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct WorklogPreview {
    week_start: String,
    week_end: String,
    work_date: Option<String>,
    entries: Vec<WorklogPreviewEntry>,
    total_minutes: i64,
    errors: Vec<String>,
}

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct WorklogTaskResult {
    issue_key: String,
    status: String,
    message: String,
}

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct WorklogSendResult {
    created: usize,
    skipped: usize,
    failed: usize,
    report_path: Option<String>,
    items: Vec<WorklogTaskResult>,
}

#[derive(serde::Deserialize)]
struct YouTrackUser {
    id: String,
}

#[derive(serde::Deserialize)]
struct ExistingDuration {
    minutes: i64,
}

#[derive(serde::Deserialize)]
struct ExistingAuthor {
    id: String,
}

#[derive(serde::Deserialize)]
struct ExistingWorkItem {
    date: i64,
    duration: ExistingDuration,
    author: ExistingAuthor,
}

#[tauri::command]
fn choose_worklog_excel() -> Option<String> {
    rfd::FileDialog::new()
        .set_title("Выберите Excel-файл с трудозатратами")
        .add_filter("Excel", &["xlsx", "xls", "xlsb"])
        .pick_file()
        .map(|path| path.to_string_lossy().into_owned())
}

#[tauri::command]
fn choose_youtrack_token_file() -> Option<String> {
    rfd::FileDialog::new()
        .set_title("Выберите файл с токеном YouTrack")
        .add_filter("Текстовый файл", &["txt", "token"])
        .pick_file()
        .map(|path| path.to_string_lossy().into_owned())
}

fn excel_serial_to_date(value: f64) -> Option<NaiveDate> {
    if !value.is_finite() { return None; }
    NaiveDate::from_ymd_opt(1899, 12, 30)?.checked_add_signed(Duration::days(value.floor() as i64))
}

fn cell_date(cell: &Data) -> Option<NaiveDate> {
    match cell {
        Data::DateTime(value) => value.as_datetime().map(|date_time| date_time.date()),
        Data::Float(value) => excel_serial_to_date(*value),
        Data::Int(value) => excel_serial_to_date(*value as f64),
        Data::String(value) => NaiveDate::parse_from_str(value.trim(), "%d.%m.%Y")
            .or_else(|_| NaiveDate::parse_from_str(value.trim(), "%Y-%m-%d"))
            .ok(),
        _ => None,
    }
}

fn is_header_row(row: &[Data]) -> bool {
    if row.len() < 7 { return false; }
    let dates = (2..7).map(|index| cell_date(&row[index])).collect::<Vec<_>>();
    dates.iter().all(Option::is_some)
        && dates.windows(2).all(|pair| pair[1].unwrap() == pair[0].unwrap() + Duration::days(1))
}

fn parse_time_cell(cell: &Data, format_regex: &Regex) -> Result<Option<i64>, String> {
    let numeric_minutes = |hours: f64| -> Result<Option<i64>, String> {
        if !hours.is_finite() { return Err("нечисловое значение".to_string()); }
        if hours <= 0.0 { return Ok(None); }
        let minutes = hours * 60.0;
        let rounded = minutes.round();
        if (minutes - rounded).abs() > 0.000001 {
            return Err("значение нельзя точно представить целым числом минут".to_string());
        }
        Ok(Some(rounded as i64))
    };
    match cell {
        Data::Empty => Ok(None),
        Data::Int(value) => numeric_minutes(*value as f64),
        Data::Float(value) => numeric_minutes(*value),
        Data::String(value) => {
            let normalized = value.trim().replace(',', ".");
            if normalized.is_empty() { return Ok(None); }
            if let Ok(hours) = normalized.parse::<f64>() { return numeric_minutes(hours); }
            if let Some(captures) = format_regex.captures(&normalized) {
                let hours = captures.get(1).map_or(0, |value| value.as_str().parse::<i64>().unwrap_or(0));
                let minutes = captures.get(2).map_or(0, |value| value.as_str().parse::<i64>().unwrap_or(0));
                let total = hours * 60 + minutes;
                return Ok((total > 0).then_some(total));
            }
            Err("неподдерживаемый формат времени".to_string())
        }
        _ => Err("неподдерживаемый тип значения".to_string()),
    }
}

fn excel_column_name(mut index: usize) -> String {
    let mut name = String::new();
    index += 1;
    while index > 0 {
        let remainder = (index - 1) % 26;
        name.insert(0, (b'A' + remainder as u8) as char);
        index = (index - 1) / 26;
    }
    name
}

#[tauri::command]
fn preview_youtrack_week(excel_path: String, selected_date: String) -> Result<WorklogPreview, String> {
    let selected = NaiveDate::parse_from_str(&selected_date, "%Y-%m-%d").map_err(|_| "Некорректная выбранная дата")?;
    let monday = selected - Duration::days(selected.weekday().num_days_from_monday() as i64);
    let friday = monday + Duration::days(4);
    let mut workbook = open_workbook_auto(excel_path.trim()).map_err(|error| format!("Не удалось открыть Excel-файл: {error}"))?;
    let range = workbook.worksheet_range("Sheet1").map_err(|error| format!("Не удалось прочитать лист Sheet1: {error}"))?;
    let rows = range.rows().collect::<Vec<_>>();
    let issue_regex = Regex::new(r"^[A-Za-z][A-Za-z0-9_]*-\d+$").map_err(|_| "Ошибка регулярного выражения")?;
    let time_regex = Regex::new(r"(?i)^(?:(\d+)h)?(?:(\d+)m)?$").map_err(|_| "Ошибка регулярного выражения")?;
    let mut totals = BTreeMap::<String, i64>::new();
    let mut errors = Vec::<String>::new();
    let mut last_day_with_time: Option<NaiveDate> = None;
    let mut matching_headers = 0usize;

    for (header_index, header) in rows.iter().enumerate() {
        if !is_header_row(header) { continue; }
        let header_dates = (2..7).map(|index| cell_date(&header[index]).unwrap()).collect::<Vec<_>>();
        if header_dates[0] != monday || header_dates[4] != friday { continue; }
        matching_headers += 1;
        for (offset, row) in rows.iter().enumerate().skip(header_index + 1) {
            if is_header_row(row) { break; }
            let issue_key = row.first().and_then(DataType::get_string).map(str::trim).filter(|value| issue_regex.is_match(value));
            for column in 2..7 {
                let Some(cell) = row.get(column) else { continue };
                match parse_time_cell(cell, &time_regex) {
                    Ok(Some(minutes)) => {
                        last_day_with_time = Some(last_day_with_time.map_or(header_dates[column - 2], |current| current.max(header_dates[column - 2])));
                        if let Some(key) = issue_key { *totals.entry(key.to_string()).or_default() += minutes; }
                    }
                    Ok(None) => {}
                    Err(message) => errors.push(format!("Sheet1!{}{}: {message}", excel_column_name(column), offset + 1)),
                }
            }
        }
    }
    if matching_headers == 0 {
        errors.push(format!("В Sheet1 не найден блок недели {}–{}", monday.format("%d.%m.%Y"), friday.format("%d.%m.%Y")));
    }
    if matching_headers > 1 {
        errors.push("В Sheet1 найдено несколько блоков выбранной недели".to_string());
    }
    let work_date = last_day_with_time.map(|date| date.format("%Y-%m-%d").to_string());
    let entries = totals.into_iter().filter(|(_, minutes)| *minutes > 0).map(|(issue_key, minutes)| WorklogPreviewEntry {
        issue_key,
        minutes,
        work_date: work_date.clone().unwrap_or_else(|| friday.format("%Y-%m-%d").to_string()),
    }).collect::<Vec<_>>();
    if matching_headers > 0 && last_day_with_time.is_none() {
        errors.push("В выбранной неделе нет положительных значений времени".to_string());
    }
    if matching_headers > 0 && entries.is_empty() {
        errors.push("В выбранной неделе нет трудозатрат по корректным ключам YouTrack".to_string());
    }
    let total_minutes = entries.iter().map(|entry| entry.minutes).sum();
    Ok(WorklogPreview {
        week_start: monday.format("%Y-%m-%d").to_string(),
        week_end: friday.format("%Y-%m-%d").to_string(),
        work_date,
        entries,
        total_minutes,
        errors,
    })
}

#[tauri::command]
fn send_youtrack_worklogs(excel_path: String, token_path: String, entries: Vec<WorklogPreviewEntry>) -> Result<WorklogSendResult, String> {
    let token = std::fs::read_to_string(token_path.trim()).map_err(|_| "Не удалось прочитать файл токена")?;
    let token = token.trim();
    if token.is_empty() { return Err("Файл токена пуст".to_string()); }
    if entries.is_empty() { return Err("Нет записей для отправки".to_string()); }
    let client = reqwest::blocking::Client::builder().build().map_err(|_| "Не удалось создать HTTP-клиент")?;
    let me = client.get(format!("{YOUTRACK_BASE_URL}/api/users/me?fields=id,login,fullName"))
        .bearer_auth(token).header("Accept", "application/json").send().map_err(|_| "Не удалось подключиться к YouTrack")?;
    if !me.status().is_success() { return Err(format!("YouTrack отклонил токен: HTTP {}", me.status().as_u16())); }
    let current_user: YouTrackUser = me.json().map_err(|_| "YouTrack вернул некорректные данные пользователя")?;
    let mut result = WorklogSendResult { created: 0, skipped: 0, failed: 0, report_path: None, items: Vec::new() };

    for entry in &entries {
        let outcome = (|| -> Result<&'static str, String> {
            let date = NaiveDate::parse_from_str(&entry.work_date, "%Y-%m-%d").map_err(|_| "некорректная дата фиксации")?;
            let timestamp = Utc.from_utc_datetime(&date.and_hms_opt(0, 0, 0).ok_or("некорректная дата")?).timestamp_millis();
            let issue = urlencoding::encode(&entry.issue_key);
            let list_url = format!("{YOUTRACK_BASE_URL}/api/issues/{issue}/timeTracking/workItems?fields=id,date,duration(minutes),author(id)&$top=-1");
            let response = client.get(list_url).bearer_auth(token).header("Accept", "application/json").send().map_err(|_| "ошибка подключения при чтении work items")?;
            if !response.status().is_success() { return Err(format!("не удалось прочитать work items: HTTP {}", response.status().as_u16())); }
            let existing: Vec<ExistingWorkItem> = response.json().map_err(|_| "некорректный ответ со списком work items")?;
            if existing.iter().any(|item| item.author.id == current_user.id && item.date == timestamp && item.duration.minutes == entry.minutes) {
                return Ok("skipped");
            }
            let create_url = format!("{YOUTRACK_BASE_URL}/api/issues/{issue}/timeTracking/workItems?fields=id,date,duration(minutes,presentation)");
            let response = client.post(create_url).bearer_auth(token).header("Accept", "application/json").header("Content-Type", "application/json")
                .json(&serde_json::json!({ "date": timestamp, "duration": { "minutes": entry.minutes } })).send().map_err(|_| "ошибка подключения при создании work item")?;
            if !response.status().is_success() { return Err(format!("создание отклонено: HTTP {}", response.status().as_u16())); }
            Ok("created")
        })();
        match outcome {
            Ok("created") => { result.created += 1; result.items.push(WorklogTaskResult { issue_key: entry.issue_key.clone(), status: "created".to_string(), message: "Создано".to_string() }); }
            Ok(_) => { result.skipped += 1; result.items.push(WorklogTaskResult { issue_key: entry.issue_key.clone(), status: "skipped".to_string(), message: "Точный дубль пропущен".to_string() }); }
            Err(message) => { result.failed += 1; result.items.push(WorklogTaskResult { issue_key: entry.issue_key.clone(), status: "failed".to_string(), message }); }
        }
    }
    if let Some(parent) = std::path::Path::new(excel_path.trim()).parent() {
        let report_name = format!("youtrack-report-{}.json", chrono::Local::now().format("%Y-%m-%d-%H%M%S"));
        let report_path = parent.join(report_name);
        if let Ok(json) = serde_json::to_string_pretty(&result) {
            if std::fs::write(&report_path, json).is_ok() { result.report_path = Some(report_path.to_string_lossy().into_owned()); }
        }
    }
    Ok(result)
}

#[derive(serde::Serialize)]
#[serde(rename_all = "camelCase")]
struct CorrespondenceFolder {
    name: String,
    path: String,
}

#[tauri::command]
fn choose_correspondence_folder() -> Option<String> {
    rfd::FileDialog::new()
        .set_title("Выберите папку переписки проекта")
        .pick_folder()
        .map(|path| path.to_string_lossy().into_owned())
}

#[tauri::command]
fn scan_correspondence(path: String) -> Result<Vec<CorrespondenceFolder>, String> {
    let root = std::path::Path::new(path.trim());
    if !root.is_dir() {
        return Err(format!("Папка переписки не найдена: {}", root.display()));
    }

    let mut folders = std::fs::read_dir(root)
        .map_err(|error| format!("Не удалось прочитать папку: {error}"))?
        .filter_map(Result::ok)
        .filter(|entry| entry.file_type().map(|kind| kind.is_dir()).unwrap_or(false))
        .map(|entry| CorrespondenceFolder {
            name: entry.file_name().to_string_lossy().into_owned(),
            path: entry.path().to_string_lossy().into_owned(),
        })
        .collect::<Vec<_>>();
    folders.sort_by(|left, right| right.name.cmp(&left.name));
    Ok(folders)
}

fn sanitize_folder_part(value: &str) -> String {
    value
        .trim()
        .chars()
        .map(|character| match character {
            '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*' => ' ',
            _ => character,
        })
        .collect::<String>()
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
}

#[tauri::command]
fn create_correspondence_draft(
    root: String,
    year: u16,
    month: u8,
    counterparty: String,
    subject: String,
) -> Result<CorrespondenceFolder, String> {
    let root_path = std::path::Path::new(root.trim());
    if !root_path.is_dir() {
        return Err("Сначала выберите существующую папку переписки".to_string());
    }
    if !(1..=12).contains(&month) {
        return Err("Месяц должен быть от 1 до 12".to_string());
    }
    let counterparty = sanitize_folder_part(&counterparty);
    let subject = sanitize_folder_part(&subject);
    if counterparty.is_empty() || subject.is_empty() {
        return Err("Укажите адресата и тему письма".to_string());
    }

    let name = format!(
        "[{year}.{month:02}.xx] [СО-{year}-xxx] ЛАБС - {counterparty}. {subject}"
    );
    let folder_path = root_path.join(&name);
    std::fs::create_dir(&folder_path)
        .map_err(|error| format!("Не удалось создать папку: {error}"))?;
    Ok(CorrespondenceFolder {
        name,
        path: folder_path.to_string_lossy().into_owned(),
    })
}

fn parse_outgoing_attachment(source: &Path) -> Result<(String, String), String> {
    let file_stem = source.file_stem().and_then(|value| value.to_str()).ok_or("Не удалось прочитать имя файла")?;
    let details = file_stem.strip_prefix("Исх. № ").ok_or("Ожидается имя вида «Исх. № СО-2026-442 от 28.08.2026.pdf»")?;
    let (number, date) = details.split_once(" от ").ok_or("В имени файла не найдены номер и дата после слова «от»")?;
    let date_parts = date.split('.').collect::<Vec<_>>();
    if date_parts.len() != 3 || date_parts[0].len() != 2 || date_parts[1].len() != 2 || date_parts[2].len() != 4 || !date_parts.iter().all(|part| part.chars().all(|character| character.is_ascii_digit())) {
        return Err("Дата в имени файла должна иметь формат дд.мм.гггг".to_string());
    }
    if !number.starts_with("СО-") {
        return Err("Номер исходящего письма должен начинаться с «СО-»".to_string());
    }
    Ok((number.to_string(), format!("{}.{}.{}", date_parts[2], date_parts[1], date_parts[0])))
}

#[tauri::command]
fn attach_outgoing_letter(folder_path: String, file_path: String) -> Result<CorrespondenceFolder, String> {
    let folder = std::path::Path::new(folder_path.trim());
    let source = std::path::Path::new(file_path.trim());
    if !folder.is_dir() {
        return Err("Папка письма не найдена".to_string());
    }
    if !source.is_file() {
        return Err("Перетащите на письмо один файл".to_string());
    }
    let (number, normalized_date) = parse_outgoing_attachment(source)?;

    let folder_name = folder.file_name().and_then(|value| value.to_str()).ok_or("Не удалось прочитать имя папки")?;
    let first_close = folder_name.find(']').ok_or("В имени папки не найден блок даты")?;
    let after_date = folder_name[first_close + 1..].trim_start();
    if !after_date.starts_with('[') {
        return Err("В имени папки не найден блок номера".to_string())
    }
    let number_close = after_date.find(']').ok_or("В имени папки не закрыт блок номера")?;
    let suffix = &after_date[number_close + 1..];
    if !suffix.trim_start().starts_with("ЛАБС -") {
        return Err("Файл можно добавить только к исходящему письму «ЛАБС - …»".to_string());
    }

    let new_name = format!("[{normalized_date}] [{number}]{}", suffix);
    let parent = folder.parent().ok_or("Не удалось определить родительскую папку")?;
    let renamed_folder = parent.join(&new_name);
    let needs_rename = renamed_folder != folder;
    if needs_rename && renamed_folder.exists() {
        return Err(format!("Папка «{new_name}» уже существует"));
    }
    if needs_rename {
        std::fs::rename(folder, &renamed_folder).map_err(|error| format!("Не удалось переименовать папку: {error}"))?;
    }

    let destination = renamed_folder.join(source.file_name().ok_or("Не удалось прочитать имя файла")?);
    if let Err(error) = std::fs::copy(source, &destination) {
        if needs_rename {
            let _ = std::fs::rename(&renamed_folder, folder);
        }
        return Err(format!("Не удалось скопировать файл в папку письма: {error}"));
    }

    Ok(CorrespondenceFolder {
        name: new_name,
        path: renamed_folder.to_string_lossy().into_owned(),
    })
}

fn validated_outlook_sources(app: &tauri::AppHandle, attachment_paths: &[String]) -> Result<Vec<PathBuf>, String> {
    if attachment_paths.is_empty() { return Err("Выберите хотя бы одно вложение".to_string()); }
    let root = app.path().app_data_dir().map_err(|error| format!("Не удалось определить папку данных: {error}"))?.join("outlook-attachments");
    let root = root.canonicalize().map_err(|_| "Локальное хранилище вложений не найдено".to_string())?;
    attachment_paths.iter().map(|value| {
        let source = Path::new(value.trim()).canonicalize().map_err(|_| "Одно из вложений больше не найдено".to_string())?;
        if !source.is_file() || !source.starts_with(&root) { return Err("Недопустимый путь к вложению".to_string()); }
        Ok(source)
    }).collect()
}

fn copy_attachment_sources(sources: &[PathBuf], destination: &Path) -> Result<(), String> {
    for source in sources {
        let file_name = source.file_name().and_then(|value| value.to_str()).ok_or("Не удалось прочитать имя вложения")?;
        let target = available_destination(destination, file_name);
        std::fs::copy(source, &target).map_err(|error| format!("Не удалось скопировать «{file_name}»: {error}"))?;
    }
    Ok(())
}

#[tauri::command]
fn file_outlook_outgoing(
    app: tauri::AppHandle,
    folder_path: String,
    attachment_paths: Vec<String>,
) -> Result<CorrespondenceFolder, String> {
    let folder = Path::new(folder_path.trim());
    if !folder.is_dir() { return Err("Выбранная папка письма не найдена".to_string()); }
    let folder_name = folder.file_name().and_then(|value| value.to_str()).ok_or("Не удалось прочитать имя папки")?;
    let outgoing_name = Regex::new(r"^\[[^\]]+\]\s+\[[^\]]+\]\s+ЛАБС\s+-").expect("valid outgoing folder regex");
    if !outgoing_name.is_match(folder_name) { return Err("Выберите папку исходящего письма".to_string()); }
    let sources = validated_outlook_sources(&app, &attachment_paths)?;
    let draft_name = Regex::new(r"(?i)^\[\d{4}\.\d{2}\.xx\]\s+\[[^\]]*xxx[^\]]*\]").expect("valid draft folder regex");
    let draft = draft_name.is_match(folder_name);
    if draft {
        let letter = sources.iter().find(|source| parse_outgoing_attachment(source).is_ok()).cloned()
            .ok_or("Для заготовки нужно выбрать файл вида «Исх. № СО-2026-442 от 28.08.2026.pdf»")?;
        let updated = attach_outgoing_letter(folder_path, letter.to_string_lossy().into_owned())?;
        let remaining = sources.into_iter().filter(|source| source != &letter).collect::<Vec<_>>();
        copy_attachment_sources(&remaining, Path::new(&updated.path))?;
        return Ok(updated);
    }
    copy_attachment_sources(&sources, folder)?;
    Ok(CorrespondenceFolder { name: folder_name.to_string(), path: folder.to_string_lossy().into_owned() })
}

#[tauri::command]
fn file_outlook_incoming(
    app: tauri::AppHandle,
    root: String,
    date: String,
    number: String,
    correspondent: String,
    subject: String,
    attachment_paths: Vec<String>,
) -> Result<CorrespondenceFolder, String> {
    let root = Path::new(root.trim());
    if !root.is_dir() { return Err("Папка переписки проекта не найдена".to_string()); }
    let parsed_date = NaiveDate::parse_from_str(date.trim(), "%Y-%m-%d").map_err(|_| "Укажите дату входящего письма".to_string())?;
    let number = sanitize_folder_part(&number);
    let correspondent = sanitize_folder_part(&correspondent);
    let subject = sanitize_folder_part(&subject);
    if number.is_empty() || correspondent.is_empty() || subject.is_empty() { return Err("Заполните номер, корреспондента и тему письма".to_string()); }
    let sources = validated_outlook_sources(&app, &attachment_paths)?;
    let name = format!("[{}] [{}] {} - ЛАБС. {}", parsed_date.format("%Y.%m.%d"), number, correspondent, subject);
    let folder = root.join(&name);
    if folder.exists() { return Err(format!("Папка «{name}» уже существует")); }
    std::fs::create_dir(&folder).map_err(|error| format!("Не удалось создать папку письма: {error}"))?;
    if let Err(error) = copy_attachment_sources(&sources, &folder) {
        let _ = std::fs::remove_dir_all(&folder);
        return Err(error);
    }
    Ok(CorrespondenceFolder { name, path: folder.to_string_lossy().into_owned() })
}

#[tauri::command]
fn rename_correspondence_folder(
    folder_path: String,
    date: String,
    number: String,
    direction: String,
    correspondent: String,
    subject: String,
) -> Result<CorrespondenceFolder, String> {
    let folder = std::path::Path::new(folder_path.trim());
    if !folder.is_dir() {
        return Err("Папка письма не найдена".to_string());
    }
    let date = sanitize_folder_part(&date);
    let number = sanitize_folder_part(&number);
    let correspondent = sanitize_folder_part(&correspondent);
    let subject = sanitize_folder_part(&subject);
    if date.is_empty() || number.is_empty() || correspondent.is_empty() || subject.is_empty() {
        return Err("Заполните дату, номер, корреспондента и тему".to_string());
    }
    let route = match direction.as_str() {
        "outgoing" => format!("ЛАБС - {correspondent}"),
        "incoming" => format!("{correspondent} - ЛАБС"),
        _ => return Err("Выберите направление письма".to_string()),
    };
    let new_name = format!("[{date}] [{number}] {route}. {subject}");
    let target = folder.parent().ok_or("Не удалось определить родительскую папку")?.join(&new_name);
    if target != folder && target.exists() {
        return Err(format!("Папка «{new_name}» уже существует"));
    }
    if target != folder {
        std::fs::rename(folder, &target).map_err(|error| format!("Не удалось переименовать папку: {error}"))?;
    }
    Ok(CorrespondenceFolder { name: new_name, path: target.to_string_lossy().into_owned() })
}

#[tauri::command]
fn delete_correspondence_folder(folder_path: String) -> Result<(), String> {
    let folder = std::path::Path::new(folder_path.trim());
    if !folder.is_dir() {
        return Err("Папка письма не найдена".to_string());
    }
    trash::delete(folder).map_err(|error| format!("Не удалось переместить папку в Корзину: {error}"))
}

#[tauri::command]
fn open_local_path(path: String) -> Result<(), String> {
    let target = std::path::Path::new(path.trim());
    if !target.exists() {
        return Err(format!("Файл или папка не найдены: {}", target.display()));
    }

    std::process::Command::new("explorer.exe")
        .arg(target)
        .spawn()
        .map(|_| ())
        .map_err(|error| format!("Не удалось открыть {}: {error}", target.display()))
}

#[tauri::command]
fn open_web_url(url: String) -> Result<(), String> {
    let target = url.trim();
    if !(target.starts_with("https://") || target.starts_with("http://")) {
        return Err("Разрешены только ссылки http:// и https://".to_string());
    }

    std::process::Command::new("explorer.exe")
        .arg(target)
        .spawn()
        .map(|_| ())
        .map_err(|error| format!("Не удалось открыть ссылку: {error}"))
}

fn show_window(app: &tauri::AppHandle, label: &str) {
    if let Some(window) = app.get_webview_window(label) {
        let _ = window.show();
        let _ = window.unminimize();
        let _ = window.set_focus();
    }
}

fn toggle_sticker(app: &tauri::AppHandle) {
    if let Some(window) = app.get_webview_window("sticker") {
        if window.is_visible().unwrap_or(false) {
            let _ = window.hide();
        } else {
            let _ = window.show();
            let _ = window.unminimize();
            let _ = window.set_focus();
        }
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_single_instance::init(|app, _args, _cwd| {
            show_window(app, "main");
        }))
        .manage(OutlookBridgeState::default())
        .invoke_handler(tauri::generate_handler![
            open_local_path,
            open_web_url,
            choose_correspondence_folder,
            scan_correspondence,
            create_correspondence_draft,
            attach_outgoing_letter,
            file_outlook_outgoing,
            file_outlook_incoming,
            rename_correspondence_folder,
            delete_correspondence_folder,
            choose_worklog_excel,
            choose_youtrack_token_file,
            preview_youtrack_week,
            send_youtrack_worklogs,
            set_outlook_bridge_key
        ])
        .plugin(
            tauri_plugin_global_shortcut::Builder::new()
                .with_handler(|app, _shortcut, event| {
                    if event.state() == ShortcutState::Pressed {
                        toggle_sticker(app);
                    }
                })
                .build(),
        )
        .setup(|app| {
            start_outlook_bridge(app.handle().clone(), app.state::<OutlookBridgeState>().inner().clone());
            if let Err(error) = app.global_shortcut().register("Ctrl+Shift+Space") {
                eprintln!("failed to register Ctrl+Shift+Space: {error}");
            }

            let open_main = MenuItem::with_id(app, "open-main", "Открыть MyPlanner", true, None::<&str>)?;
            let show_sticker = MenuItem::with_id(app, "show-sticker", "Показать стикер", true, None::<&str>)?;
            let separator = PredefinedMenuItem::separator(app)?;
            let quit = MenuItem::with_id(app, "quit", "Выход", true, None::<&str>)?;
            let menu = Menu::with_items(app, &[&open_main, &show_sticker, &separator, &quit])?;

            TrayIconBuilder::new()
                .icon(app.default_window_icon().expect("application icon is missing").clone())
                .tooltip("MyPlanner")
                .menu(&menu)
                .show_menu_on_left_click(false)
                .on_menu_event(|app, event| match event.id.as_ref() {
                    "open-main" => show_window(app, "main"),
                    "show-sticker" => show_window(app, "sticker"),
                    "quit" => app.exit(0),
                    _ => {}
                })
                .on_tray_icon_event(|tray, event| {
                    if let TrayIconEvent::Click {
                        button: MouseButton::Left,
                        button_state: MouseButtonState::Up,
                        ..
                    } = event
                    {
                        show_window(tray.app_handle(), "main");
                    }
                })
                .build(app)?;

            Ok(())
        })
        .on_window_event(|window, event| {
            if let tauri::WindowEvent::CloseRequested { api, .. } = event {
                api.prevent_close();
                let _ = window.hide();
            }
        })
        .run(tauri::generate_context!())
        .expect("error while running MyPlanner");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_supported_worklog_formats() {
        let expression = Regex::new(r"(?i)^(?:(\d+)h)?(?:(\d+)m)?$").unwrap();
        assert_eq!(parse_time_cell(&Data::Float(2.0), &expression).unwrap(), Some(120));
        assert_eq!(parse_time_cell(&Data::Float(0.5), &expression).unwrap(), Some(30));
        assert_eq!(parse_time_cell(&Data::Float(1.25), &expression).unwrap(), Some(75));
        assert_eq!(parse_time_cell(&Data::String("2h25m".into()), &expression).unwrap(), Some(145));
        assert_eq!(parse_time_cell(&Data::String("2h".into()), &expression).unwrap(), Some(120));
        assert_eq!(parse_time_cell(&Data::String("45m".into()), &expression).unwrap(), Some(45));
        assert_eq!(parse_time_cell(&Data::Float(0.0), &expression).unwrap(), None);
        assert_eq!(parse_time_cell(&Data::Float(-1.0), &expression).unwrap(), None);
        assert!(parse_time_cell(&Data::Float(0.001), &expression).is_err());
    }

    #[test]
    fn converts_excel_serial_dates() {
        assert_eq!(excel_serial_to_date(44543.0).unwrap(), NaiveDate::from_ymd_opt(2021, 12, 13).unwrap());
    }
}
