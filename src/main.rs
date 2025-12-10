use std::{num::NonZeroUsize, path::PathBuf, sync::Arc};

use iced::{
    Element, Task, Theme,
    theme::Palette,
    widget::{
        Button, Container, Text, column, container, row,
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
    sheet_name: String,
    title_row_index: NonZeroUsize,
}

#[derive(Clone)]
enum Message {
    CardTitleChanged(String),
    ExcelPathChanged(PathBuf),
    SheetsNamesLoaded(Arc<[String]>),
    SheetNameSelected(usize),
    TitlRowPossibleIndexsLoaded(Arc<[usize]>),
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
            sheet_name: Default::default(),
            title_row_index: NonZeroUsize::new(1).unwrap(),
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
            }
            Message::SheetsNamesLoaded(items) => todo!(),
            Message::SheetNameSelected(_) => todo!(),
            Message::TitlRowPossibleIndexsLoaded(items) => todo!(),
            Message::TitlRowIndexSelected(_) => todo!(),
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
        let txt = "عنوان الكارت";
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
        self.card_title_view()
    }
    fn title_row_index_view(&self) -> Element<'_, Message> {
        self.card_title_view()
    }
    fn submit_button_view(&self) -> Element<'_, Message> {
        self.card_title_view()
    }
}
