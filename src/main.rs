use std::{env::home_dir, fs::File, io::Write, path::PathBuf, sync::Arc};

use crate::web_render::web_cards;
use calamine::{DeError, RangeDeserializerBuilder, Reader, Xlsx, open_workbook};
use iced::{
    Alignment, Element, Task, Theme,
    theme::Palette,
    widget::{
        Button, Container, PickList, Row, Text, checkbox, column, container, row,
        text_input::{self, Style},
    },
};

mod web_render;

fn main() {
    iced::application(App::new, App::update, App::view)
        .run()
        .unwrap();
}

struct App {
    card_title: String,
    excel_path: PathBuf,
    exel_path_exists: bool,
    exel_path_is_excel: bool,
    all_sheets_names: Arc<[String]>,
    sheet_name: Option<String>,
    all_rows_indexes: Arc<[usize]>,
    title_row_index: Option<usize>,
    all_titles_names: Vec<(bool, String)>,
    rendered_at: Option<PathBuf>,
}

#[derive(Clone)]
enum Message {
    CardTitleChanged(String),
    ExcelPathChanged(PathBuf),
    SheetNameSelected(String),
    TitlRowIndexSelected(usize),
    PickExelFile,
    ToggleTitle((usize, bool)),
    Render,
}

impl App {
    fn new() -> Self {
        Self {
            card_title: Default::default(),
            excel_path: Default::default(),
            exel_path_exists: false,
            exel_path_is_excel: false,
            all_sheets_names: Arc::new([]),
            sheet_name: Default::default(),
            all_rows_indexes: Arc::new([]),
            title_row_index: None,
            all_titles_names: Vec::new(),
            rendered_at: None,
        }
    }

    fn update(&mut self, message: Message) -> Task<Message> {
        match message {
            Message::CardTitleChanged(title) => {
                self.card_title = title;
            }
            Message::ExcelPathChanged(path_buf) => {
                self.exel_path_exists = path_buf.exists();
                self.exel_path_is_excel = path_buf
                    .extension()
                    .is_some_and(|x| ["xls", "xlsx", "xlsb", "ods"].contains(&x.to_str().unwrap()));
                self.excel_path = path_buf;
                if self.exel_path_exists && self.exel_path_is_excel {
                    match open_workbook::<Xlsx<_>, _>(&self.excel_path) {
                        Ok(wb) => {
                            self.all_sheets_names = wb.sheet_names().into();
                        }
                        Err(err) => {
                            eprintln!("Error : could not open workbook due to -> {err}");
                        }
                    };
                }
            }
            Message::SheetNameSelected(sheet) => {
                match rows_range(&self.excel_path, &sheet) {
                    Ok((top, bottom)) => {
                        self.all_rows_indexes = ((top + 1)..=bottom).collect();
                        self.sheet_name = Some(sheet);
                    }
                    Err(err) => {
                        eprintln!("Error : could not get rows range due to -> {err}");
                    }
                };
            }
            Message::TitlRowIndexSelected(index) => {
                let Some(sheet_name) = &self.sheet_name else {
                    return Task::none();
                };
                match get_titles(&self.excel_path, sheet_name, index - 1) {
                    Ok(titles) => {
                        self.all_titles_names = titles.into_iter().map(|x| (false, x)).collect();
                        self.title_row_index = Some(index);
                    }
                    Err(err) => {
                        eprintln!("Error : could not fetch titles row due to -> {err}");
                    }
                };
            }
            Message::PickExelFile => todo!(),
            Message::ToggleTitle((index, exists)) => {
                self.all_titles_names[index].0 = exists;
            }
            Message::Render => {
                let (Some(title_row_index), Some(sheet_name)) =
                    (self.title_row_index, &self.sheet_name)
                else {
                    return Task::none();
                };
                let html = web_cards(
                    self.card_title.clone(),
                    title_row_index,
                    &self.excel_path,
                    sheet_name,
                    self.all_titles_names
                        .iter()
                        .enumerate()
                        .filter(|x| x.1.0)
                        .map(|x| x.0)
                        .collect(),
                );
                let path = home_dir().unwrap().join("kvg_index.html");
                let mut file = File::create(&path).unwrap();
                file.write_all(&html.into_bytes()).unwrap();
                self.rendered_at = Some(path);
            }
        }
        Task::none()
    }

    fn view(&self) -> Element<'_, Message> {
        let ct = self.card_title_view();
        let et = self.excel_path_view();
        let sn = self.sheet_name_view();
        let tri = self.title_row_index_view();
        let sb = self.submit_button_view();
        let trp = self.titles_row_pick_view();
        let col = column![ct, et, sn, tri, trp, sb].spacing(5.).padding(5.);
        Container::new(col).style(container::rounded_box).into()
    }

    fn card_title_view(&self) -> Element<'_, Message> {
        let txt = "عنوان الكارت";
        let text = Text::new(txt);
        let input =
            text_input::TextInput::new(txt, &self.card_title).on_input(Message::CardTitleChanged);
        let row = row![text, input];
        Container::new(row).style(container::rounded_box).into()
    }
    fn excel_path_view(&self) -> Element<'_, Message> {
        let txt = "موقع ملف الاكسل";
        let text = Text::new(txt);
        let input = text_input::TextInput::new(txt, &self.excel_path.display().to_string())
            .on_input(|x| {
                let Ok(x) = x.parse();
                Message::ExcelPathChanged(x)
            })
            .style(|th: &Theme, _| {
                let Palette {
                    success,
                    danger,
                    warning,
                    primary,
                    ..
                } = th.palette();
                let color = if self.exel_path_exists && self.exel_path_is_excel {
                    success
                } else if self.exel_path_exists {
                    warning
                } else {
                    danger
                };
                Style {
                    value: color,
                    background: iced::Background::Color(iced::Color::WHITE),
                    border: iced::Border {
                        color,
                        width: 3.,
                        radius: iced::border::Radius::new(5.),
                    },
                    icon: color,
                    placeholder: color,
                    selection: primary,
                }
            });
        let browse = Button::new("Browse").on_press(Message::PickExelFile);
        let input = row![input, browse].spacing(3.).padding(3.);
        let row = row![text, input];
        Container::new(row).style(container::rounded_box).into()
    }
    fn sheet_name_view(&self) -> Element<'_, Message> {
        let txt = "اسم الشييت";
        let text = Text::new(txt);
        let input = PickList::new(
            self.all_sheets_names.clone(),
            self.sheet_name.clone(),
            Message::SheetNameSelected,
        );
        let row = row![text, input];
        Container::new(row).style(container::rounded_box).into()
    }
    fn title_row_index_view(&self) -> Element<'_, Message> {
        let txt = "مسلسل صف العناوين";
        let text = Text::new(txt);
        let input = PickList::new(
            self.all_rows_indexes.clone(),
            self.title_row_index,
            Message::TitlRowIndexSelected,
        );
        let row = row![text, input];
        Container::new(row).style(container::rounded_box).into()
    }
    fn titles_row_pick_view(&self) -> Element<'_, Message> {
        let txt = "اختار الاعمدة";
        let text = Text::new(txt);
        let row = self.all_titles_names.iter().enumerate().fold(
            Row::new(),
            |acc, (index, (exists, title))| {
                acc.push(
                    checkbox(*exists)
                        .label(title)
                        .on_toggle(move |ch| Message::ToggleTitle((index, ch))),
                )
            },
        );
        let col = column![text, row.wrap()];
        Container::new(col).style(container::rounded_box).into()
    }
    fn submit_button_view(&self) -> Element<'_, Message> {
        let clickable = self.all_titles_names.iter().filter(|x| x.0).count() > 0;
        let b =
            Button::new(if clickable { "تمام" } else { "افندم!" }).on_press_maybe(if clickable {
                Some(Message::Render)
            } else {
                None
            });
        let path = self
            .rendered_at
            .as_ref()
            .map(|x| Text::new(format!("rendered at : {}", x.display())));
        column![b, path].align_x(Alignment::Center).into()
    }
}

fn rows_range(path: &PathBuf, sheetname: &str) -> Result<(usize, usize), calamine::Error> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook.worksheet_range(sheetname)?;
    let top = range.start().map(|x| x.0);
    let bottom = range.end().map(|x| x.0);
    match (top, bottom) {
        (Some(t), Some(b)) => Ok((t as usize, b as usize)),
        _ => Ok((0, 0)),
    }
}

fn get_titles(
    path: &PathBuf,
    sheetname: &str,
    headers_index: usize,
) -> Result<Vec<String>, calamine::Error> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook.worksheet_range(sheetname)?;

    let mut iter = RangeDeserializerBuilder::new()
        .has_headers(false)
        .from_range(&range)?;

    let headers: Vec<String> = iter
        .nth(headers_index)
        .unwrap_or(Err(DeError::HeaderNotFound(format!(
            "Error number {headers_index} should contain headers"
        ))))?;

    Ok(headers)
}
