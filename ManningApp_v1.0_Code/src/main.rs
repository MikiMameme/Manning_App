//
// ManningApp 1.0
// Created by K.N (2026)
// Developed with the assistance of AI (Google Antigravity, Claude, Gemini)
// License: MIT
//

#![windows_subsystem = "windows"]

use eframe::egui;
use chrono::{FixedOffset, Utc, Datelike, NaiveDate};
use calamine::{open_workbook, Reader, Xlsx, DataType};
use std::path::PathBuf;

fn main() -> eframe::Result<()> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([750.0, 700.0])
            .with_min_inner_size([600.0, 500.0]),
        ..Default::default()
    };
    eframe::run_native(
        "勤務表ビューア",
        options,
        Box::new(|cc| {
            let mut fonts = egui::FontDefinitions::default();
            fonts.font_data.insert(
                "my_font".to_owned(),
                egui::FontData::from_static(include_bytes!("../NotoSansJP-Regular.ttf")),
            );
            fonts.families.get_mut(&egui::FontFamily::Proportional).unwrap()
                .insert(0, "my_font".to_owned());
            fonts.families.get_mut(&egui::FontFamily::Monospace).unwrap()
                .insert(0, "my_font".to_owned());
            cc.egui_ctx.set_fonts(fonts);

            Box::new(ManningApp::default())
        }),
    )
}

struct ManningApp {
    date_display: String,
    today: NaiveDate,
    staff_names: Vec<String>,
    schedule_text: String,
    status_message: String,
}

impl Default for ManningApp {
    fn default() -> Self {
        let offset = FixedOffset::east_opt(9 * 3600).unwrap();
        let now = Utc::now().with_timezone(&offset);
        let weekdays = ["日", "月", "火", "水", "木", "金", "土"];
        let weekday_str = weekdays[now.weekday().num_days_from_sunday() as usize];
        let date_str = format!("{}月{}日({})", now.month(), now.day(), weekday_str);
        let today = now.date_naive();

        Self {
            date_display: date_str,
            today,
            staff_names: vec!["".to_string(); 4],
            schedule_text: String::new(),
            status_message: String::new(),
        }
    }
}

impl ManningApp {
    /// エクセル勤務表を読み込んで、本日のシフトを自動入力する
    fn load_excel(&mut self, path: PathBuf) {
        let result: Result<Xlsx<_>, _> = open_workbook(&path);
        match result {
            Ok(mut workbook) => {
                // 最初のシートを使用
                let sheets = workbook.sheet_names().to_vec();
                if sheets.is_empty() {
                    self.status_message = "❌ シートが見つかりません".to_string();
                    return;
                }
                let sheet_name = sheets[0].clone();

                if let Ok(range) = workbook.worksheet_range(&sheet_name) {
                    let mut found_date_pos: Option<(usize, usize)> = None; // (row, col)
                    let today_day = self.today.day();
                    let today_month = self.today.month();

                    // 1. 全行・全列を走査して「本日の日付」セルを探す
                    'outer: for (row_idx, row) in range.rows().enumerate() {
                        for (col_idx, cell) in row.iter().enumerate() {
                            // 1. シリアル値または日付型として解釈 (features=["dates"]が必要)
                            if let Some(dt) = cell.as_date() {
                                if dt.month() == today_month && dt.day() == today_day {
                                    found_date_pos = Some((row_idx, col_idx));
                                    break 'outer;
                                }
                            }

                            // 2. 数値として解釈 (日にちのみ)
                            if let Some(day_i64) = cell.as_i64() {
                                let day = day_i64 as u32;
                                if day == today_day {
                                    found_date_pos = Some((row_idx, col_idx));
                                    break 'outer;
                                }
                            }

                            // 3. 文字列として処理
                            let cell_str = format!("{}", cell);
                            let cell_str_norm = to_hankaku(&cell_str);

                             // 正規化後の文字列で数値パース試行 (例: "１４" -> "14")
                            if let Ok(day) = cell_str_norm.trim().parse::<u32>() {
                                if day == today_day {
                                    found_date_pos = Some((row_idx, col_idx));
                                    break 'outer;
                                }
                            }

                            // 文字列日付パース (yyyy/m/d, m/d)
                            let parts: Vec<&str> = cell_str_norm.split('/').collect();
                            if parts.len() >= 2 {
                                // m/d または yyyy/m/d
                                let m_idx = if parts.len() == 2 { 0 } else { 1 };
                                let d_idx = if parts.len() == 2 { 1 } else { 2 };
                                
                                if let (Ok(m), Ok(d)) = (parts[m_idx].trim().parse::<u32>(), parts[d_idx].trim().parse::<u32>()) {
                                    if m == today_month && d == today_day {
                                        found_date_pos = Some((row_idx, col_idx));
                                        break 'outer;
                                    }
                                }
                            }

                            // 日付形式チェック (contains)
                            if cell_str_norm.contains(&format!("{}/{}", today_month, today_day))
                                || cell_str_norm.contains(&format!("{}月{}日", today_month, today_day)) {
                                found_date_pos = Some((row_idx, col_idx));
                                break 'outer;
                            }
                        }
                    }

                    if let Some((date_row_idx, shift_col_idx)) = found_date_pos {
                         // シフト種別: 早番, 日勤, 遅番, 夜勤
                        let shift_keywords = ["早", "日", "遅", "夜"];
                        let mut shift_staff: Vec<Vec<String>> = vec![vec![]; 4];

                        // 2. 日付行の2つ下（曜日行の下）からデータ読み込み開始
                        // shift_col_idx 列がシフト値。名前はそれより左の列から探す。
                        for row in range.rows().skip(date_row_idx + 2) {
                            if row.len() <= shift_col_idx { continue; }
                            
                            let shift_val = format!("{}", row[shift_col_idx]).trim().to_string();
                            if shift_val.is_empty() { continue; }

                            // 名前列の探索（シフト列より左にある非空セルを採用: 左端優先）
                            let mut name = String::new();
                            for col in 0..shift_col_idx {
                                if let Some(cell) = row.get(col) {
                                    let val = format!("{}", cell).trim().to_string();
                                    if !val.is_empty() {
                                        name = val;
                                        break;
                                    }
                                }
                            }

                            if name.is_empty() { continue; }

                            for (i, keyword) in shift_keywords.iter().enumerate() {
                                if shift_val.contains(keyword) {
                                    shift_staff[i].push(name.clone());
                                    break;
                                }
                            }
                        }

                        // シフト欄に入力
                        for i in 0..4 {
                            if !shift_staff[i].is_empty() {
                                self.staff_names[i] = shift_staff[i].join("、");
                            }
                        }
                         self.status_message = format!("✅ エクセルを読み込みました (行:{}, 列:{})", date_row_idx + 1, shift_col_idx + 1);
                    } else {
                        self.status_message = format!("❌ 本日({}月{}日)の日付列が見つかりません", today_month, today_day);
                    }

                } else {
                    self.status_message = "❌ シートの読み込みに失敗しました".to_string();
                }
            }
            Err(e) => {
                self.status_message = format!("❌ ファイルを開けません: {}", e);
            }
        }
    }
}

impl eframe::App for ManningApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {

            // === 上部バー: 日付 + ボタン ===
            ui.horizontal(|ui| {
                ui.heading(format!("📅 {}", self.date_display));
                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    if ui.button(egui::RichText::new("スクショ\n印刷").size(14.0)).clicked() {
                        ctx.send_viewport_cmd(egui::ViewportCommand::Screenshot);
                    }
                    if ui.button(egui::RichText::new("エクセル\n読込み").size(14.0)).clicked() {
                        if let Some(path) = rfd::FileDialog::new()
                            .add_filter("Excel", &["xlsx", "xls"])
                            .pick_file() {
                            self.load_excel(path);
                        }
                    }
                });
            });
            ui.add_space(5.0);

            // === シフト進捗バー (削除済み) ===
            let empty_count = self.staff_names.iter().filter(|n| n.is_empty()).count();
            
            ui.add_space(10.0);

            // === シフト入力欄（左） + テーブル表示（右） ===
            let shift_labels = ["早番", "日勤", "遅番", "夜勤"];

            ui.horizontal(|ui| {
                // 左側: シフト入力欄
                ui.vertical(|ui| {
                    ui.set_min_width(200.0);
                    for (i, label) in shift_labels.iter().enumerate() {
                        ui.horizontal(|ui| {
                            ui.label(egui::RichText::new(format!("{}:", label)).size(16.0).strong());
                            ui.add_sized(
                                [120.0, 24.0],
                                egui::TextEdit::singleline(&mut self.staff_names[i])
                                    .font(egui::TextStyle::Body),
                            );
                        });
                        ui.add_space(2.0);
                    }
                });

                ui.add_space(20.0);

                // 右側: テーブル形式シフト表
                egui::Frame::none()
                    .stroke(egui::Stroke::new(2.0, egui::Color32::BLACK))
                    .inner_margin(0.0)
                    .show(ui, |ui| {
                        egui::Grid::new("shift_table")
                            .striped(false)
                            .min_col_width(70.0)
                            .show(ui, |ui| {
                                // ヘッダー行
                                for label in &shift_labels {
                                    egui::Frame::none()
                                        .stroke(egui::Stroke::new(1.0, egui::Color32::BLACK))
                                        .inner_margin(8.0)
                                        .show(ui, |ui| {
                                            ui.label(egui::RichText::new(*label).strong().size(18.0));
                                        });
                                }
                                ui.end_row();

                                // スタッフ名行
                                for name in &self.staff_names {
                                    egui::Frame::none()
                                        .stroke(egui::Stroke::new(1.0, egui::Color32::BLACK))
                                        .inner_margin(8.0)
                                        .show(ui, |ui| {
                                            let display = if name.is_empty() { "―" } else { name.as_str() };
                                            ui.label(egui::RichText::new(display).size(18.0));
                                        });
                                }
                                ui.end_row();
                            });
                    });
            });

            ui.add_space(10.0);

            // === ステータスメッセージ ===
            if empty_count > 0 {
                ui.label(egui::RichText::new(format!("❌ シフトに不備があります（あと{}名未配置）", empty_count))
                    .color(egui::Color32::RED)
                    .strong());
            } else {
                ui.label(egui::RichText::new("✅ 今日の配置はOKです！")
                    .color(egui::Color32::GREEN)
                    .strong());
            }

            // エクセル読み込み結果メッセージ
            if !self.status_message.is_empty() {
                ui.label(egui::RichText::new(&self.status_message).size(12.0).italics());
            }

            ui.add_space(10.0);

            // === ★本日の予定セクション ===
            egui::Frame::none()
                .stroke(egui::Stroke::new(2.0, egui::Color32::BLACK))
                .inner_margin(15.0)
                .rounding(5.0)
                .show(ui, |ui| {
                    ui.set_width(ui.available_width());
                    ui.heading(egui::RichText::new("★本日の予定").size(22.0));
                    ui.add_space(10.0);

                    let available_height = ui.available_height().max(200.0);
                    ui.add_sized(
                        [ui.available_width(), available_height - 30.0],
                        egui::TextEdit::multiline(&mut self.schedule_text)
                            .font(egui::FontId::proportional(18.0))
                            .frame(false)
                            .desired_width(f32::INFINITY),
                    );
                });
        });

        // === スクリーンショット処理 ===
        if let Some(screenshot) = ctx.input(|i| i.raw.events.iter().find_map(|e| {
            if let egui::Event::Screenshot { image, .. } = e {
                Some(image.clone())
            } else {
                None
            }
        })) {
            if let Some(path) = rfd::FileDialog::new()
                .add_filter("PNG", &["png"])
                .save_file() {

                let pixels: Vec<u8> = screenshot.pixels.iter().flat_map(|p| {
                    [p.r(), p.g(), p.b(), p.a()]
                }).collect();

                if let Err(err) = image::save_buffer(
                    path,
                    &pixels,
                    screenshot.width() as u32,
                    screenshot.height() as u32,
                    image::ColorType::Rgba8,
                ) {
                    eprintln!("Error saving screenshot: {}", err);
                }
            }
        }
    }
}

/// 全角英数字・スペースを半角に変換するヘルパー関数
fn to_hankaku(s: &str) -> String {
    s.chars()
        .map(|c| match c {
            '０'..='９' => char::from_u32(c as u32 - '０' as u32 + '0' as u32).unwrap(),
            'Ａ'..='Ｚ' => char::from_u32(c as u32 - 'Ａ' as u32 + 'A' as u32).unwrap(),
            'ａ'..='ｚ' => char::from_u32(c as u32 - 'ａ' as u32 + 'a' as u32).unwrap(),
            '　' => ' ',
            _ => c,
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::to_hankaku;

    #[test]
    fn test_to_hankaku_conv() {
        assert_eq!(to_hankaku("１２３"), "123");
        assert_eq!(to_hankaku("　"), " ");
        assert_eq!(to_hankaku("ＡＢＣ"), "ABC");
        assert_eq!(to_hankaku("ａｂｃ"), "abc");
        assert_eq!(to_hankaku("２月１４日"), "2月14日");
        assert_eq!(to_hankaku("12月25日"), "12月25日");
    }
}