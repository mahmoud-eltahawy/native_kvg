use crate::web_render::web_cards;
use calamine::{DeError, RangeDeserializerBuilder, Reader, Xlsx, open_workbook};
use iced::{
    Alignment, Background, Element, Length, Shadow, Task, Theme,
    border::Radius,
    overlay::menu,
    theme::Palette,
    widget::{
        Button, Container, PickList, Row, Scrollable, Text, checkbox, column, container, row,
        text_input::{Style, TextInput},
    },
};
use rfd::FileDialog;
use std::{
    env::home_dir,
    fs::{self, File},
    io::{self, Write},
    path::PathBuf,
    sync::Arc,
};

mod web_render;

const XLSX_FILTERS: [&str; 4] = ["xls", "xlsx", "xlsb", "ods"];

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
                    .is_some_and(|x| XLSX_FILTERS.contains(&x.to_str().unwrap()));
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
                } else {
                    self.all_sheets_names = Arc::new([]);
                    self.sheet_name = None;
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
            Message::PickExelFile => {
                if let Some(path) = pick_file() {
                    self.excel_path = path;
                }
            }
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
        let col = column![ct, et, sn, tri, trp, sb]
            .spacing(25.)
            .padding(5.)
            .align_x(Alignment::Center);
        let col = Scrollable::new(col);
        Container::new(col)
            .height(Length::Fill)
            .padding(80.)
            .align_x(Alignment::Center)
            .align_y(Alignment::Center)
            .style(|theme: &Theme| {
                let Palette {
                    background,
                    text,
                    primary,
                    ..
                } = theme.palette();
                container::Style {
                    text_color: Some(text),
                    background: Some(Background::Color(background)),
                    border: iced::Border {
                        color: primary,
                        width: 7.,
                        radius: Radius::new(50.),
                    },
                    shadow: Shadow::default(),
                    snap: true,
                }
            })
            .into()
    }

    fn card_title_view(&self) -> Element<'_, Message> {
        let txt = "عنوان الكارت";
        let text = (!self.card_title.is_empty()).then_some(Text::new(txt).size(20.));
        let input = TextInput::new(txt, &self.card_title)
            .align_x(Alignment::Center)
            .padding(20.)
            .size(25.)
            .style(|theme: &Theme, _| {
                let Palette {
                    background,
                    text,
                    primary,
                    warning,
                    danger,
                    ..
                } = theme.palette();
                Style {
                    background: Background::Color(background),
                    border: iced::Border {
                        color: if self.card_title.is_empty() {
                            danger
                        } else {
                            primary
                        },
                        width: 3.,
                        radius: Radius::new(20.),
                    },
                    icon: text,
                    placeholder: warning,
                    value: text,
                    selection: danger,
                }
            })
            .on_input(Message::CardTitleChanged);
        row![input, text]
            .align_y(Alignment::Center)
            .padding(20.)
            .spacing(20.)
            .into()
    }
    fn excel_path_view(&self) -> Element<'_, Message> {
        let txt = "موقع ملف الاكسل";
        let text = (self.excel_path != PathBuf::new()).then_some(Text::new(txt).size(20.));
        let input = TextInput::new(txt, &self.excel_path.display().to_string())
            .align_x(Alignment::Center)
            .padding(10.)
            .size(25.)
            .style(|theme: &Theme, _| {
                let Palette {
                    background,
                    text,
                    primary,
                    warning,
                    danger,
                    ..
                } = theme.palette();
                Style {
                    background: Background::Color(background),
                    border: iced::Border {
                        color: primary,
                        width: 3.,
                        radius: Radius::new(20.),
                    },
                    icon: text,
                    placeholder: warning,
                    value: text,
                    selection: danger,
                }
            })
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
                        radius: Radius::new(5.),
                    },
                    icon: color,
                    placeholder: color,
                    selection: primary,
                }
            });
        let browse = Button::new("اختيار ملف").on_press(Message::PickExelFile);

        let ac = path_autocomplete(&self.excel_path).ok().map(|paths| {
            paths
                .into_iter()
                .fold(Row::new(), |acc, path| {
                    acc.push(
                        Button::new(Text::new(path.display().to_string()))
                            .on_press(Message::ExcelPathChanged(path)),
                    )
                })
                .spacing(5.)
                .wrap()
        });
        let input = row![browse, input]
            .spacing(10.)
            .spacing(3.)
            .padding(3.)
            .align_y(Alignment::Center);
        let row = row![input, text].spacing(10.).align_y(Alignment::Center);
        column![row, ac].into()
    }
    fn sheet_name_view(&self) -> Element<'_, Message> {
        let txt = "اسم الشييت";
        let text = self.sheet_name.as_ref().map(|_| Text::new(txt));
        let input = PickList::new(
            self.all_sheets_names.clone(),
            self.sheet_name.clone(),
            Message::SheetNameSelected,
        )
        .menu_style(|theme: &Theme| {
            let Palette {
                background,
                text,
                primary,
                success,
                ..
            } = theme.palette();
            //
            menu::Style {
                border: iced::Border {
                    color: primary,
                    width: 3.,
                    radius: Radius::new(3.),
                },
                background: Background::Color(background),
                text_color: text,
                selected_text_color: success,
                selected_background: Background::Color(background),
                shadow: Shadow::default(),
            }
        })
        .text_size(20.)
        .padding(10.)
        .placeholder(txt);
        row![input, text]
            .align_y(Alignment::Center)
            .spacing(20.)
            .into()
    }
    fn title_row_index_view(&self) -> Element<'_, Message> {
        let txt = "مسلسل صف العناوين";
        let text = self
            .title_row_index
            .as_ref()
            .map(|_| Text::new(txt).align_y(Alignment::Center));
        let input = PickList::new(
            self.all_rows_indexes.clone(),
            self.title_row_index,
            Message::TitlRowIndexSelected,
        )
        .text_size(20.)
        .placeholder(txt);
        row![input, text]
            .spacing(15.)
            .align_y(Alignment::Center)
            .into()
    }
    fn titles_row_pick_view(&self) -> Element<'_, Message> {
        let txt = "اختر الاعمدة";
        let text = (!self.all_titles_names.is_empty()).then_some(Text::new(txt));
        let titles_row = self
            .all_titles_names
            .iter()
            .enumerate()
            .fold(Row::new(), |acc, (index, (exists, title))| {
                acc.push(
                    checkbox(*exists)
                        .size(20.)
                        .text_size(20.)
                        .label(title)
                        .spacing(20.)
                        .on_toggle(move |ch| Message::ToggleTitle((index, ch))),
                )
            })
            .spacing(20.)
            .padding(5.)
            .align_y(Alignment::Center);
        column![text, titles_row.wrap()].spacing(20.).into()
    }
    fn submit_button_view(&self) -> Element<'_, Message> {
        let clickable =
            self.all_titles_names.iter().filter(|x| x.0).count() > 0 && !self.card_title.is_empty();
        let submit = Button::new(if clickable { "تمام" } else { "افندم!" })
            .on_press_maybe(if clickable {
                Some(Message::Render)
            } else {
                None
            })
            .padding(20.);
        let rendered_at = self
            .rendered_at
            .as_ref()
            .map(|x| Text::new(format!("rendered at : {}", x.display())));
        column![submit, rendered_at]
            .align_x(Alignment::Center)
            .into()
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

fn pick_file() -> Option<PathBuf> {
    FileDialog::new().pick_file()
}

fn path_autocomplete(path: &PathBuf) -> Result<Vec<PathBuf>, io::Error> {
    let mut paths = if path.exists() {
        let mut enteries = fs::read_dir(path)?;
        let mut paths = Vec::new();
        while let Some(entry) = enteries.next().transpose()? {
            paths.push(entry.path());
        }
        paths
    } else if path.parent().is_some_and(|x| x.exists()) {
        let parent = path.parent().unwrap();
        let name = path.file_name().unwrap().to_str().unwrap().to_lowercase();
        let mut enteries = fs::read_dir(parent)?;
        let mut paths = Vec::new();
        while let Some(entry) = enteries.next().transpose()? {
            let epath = entry.path();
            if epath
                .file_name()
                .and_then(|x| x.to_str())
                .is_some_and(|x| x.to_lowercase().starts_with(&name))
            {
                paths.push(epath);
            }
        }
        paths
    } else {
        Vec::new()
    };
    paths.sort();
    Ok(paths)
}
