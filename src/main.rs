use std::{path::PathBuf, sync::Arc};

use calamine::{Reader, Xlsx, open_workbook};
use iced::{
    Element, Task, Theme,
    theme::Palette,
    widget::{
        Button, Container, PickList, Text, column, container, row,
        text_input::{self, Style},
    },
};

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
}

#[derive(Clone)]
enum Message {
    CardTitleChanged(String),
    ExcelPathChanged(PathBuf),
    SheetNameSelected(String),
    TitlRowIndexSelected(usize),
    PickExelFile,
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
                        self.all_rows_indexes = (top..bottom).collect();
                        self.sheet_name = Some(sheet);
                    }
                    Err(err) => {
                        eprintln!("Error : could not get rows range due to -> {err}");
                    }
                };
            }
            Message::TitlRowIndexSelected(index) => {
                self.title_row_index = Some(index);
            }
            Message::PickExelFile => todo!(),
        }
        Task::none()
    }

    fn view(&self) -> Element<'_, Message> {
        let ct = self.card_title_view();
        let et = self.excel_path_view();
        let sn = self.sheet_name_view();
        let tri = self.title_row_index_view();
        let sb = self.submit_button_view();
        let col = column![ct, et, sn, tri, sb].spacing(5.).padding(5.);
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
                    ..
                } = th.palette();
                let c = if self.exel_path_exists && self.exel_path_is_excel {
                    success
                } else if self.exel_path_exists {
                    warning
                } else {
                    danger
                };
                Style {
                    value: c,
                    background: iced::Background::Color(iced::Color::WHITE),
                    border: iced::Border {
                        color: c,
                        width: 3.,
                        radius: iced::border::Radius::new(5.),
                    },
                    icon: c,
                    placeholder: c,
                    selection: c,
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
    fn submit_button_view(&self) -> Element<'_, Message> {
        self.card_title_view()
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
