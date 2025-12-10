use leptos::{either::Either, prelude::*};
use std::path::PathBuf;

const CSS: &str = include_str!("../index.css");

pub fn web_cards(
    title: String,
    title_row_index: usize,
    path: &PathBuf,
    sheet: &str,
    columns_indexs: Vec<usize>,
) -> String {
    let cards = get_cards(title_row_index, path, sheet, columns_indexs);
    let cards = match cards {
        Ok(cards) => Either::Left(view! {<Cards cards title/>}),
        Err(err) => Either::Right(view! {
            <h3>something bad happend</h3>
            <p>{err.to_string()}</p>
        }),
    };
    view! {
        <!DOCTYPE html>
        <html dir="rtl" lang="ar">
            <head>
                <meta charset="utf-8"/>
                <meta name="viewport" content="width=device-width, initial-scale=1"/>
                <title>kvg</title>
                <style>{CSS}</style>
            </head>
            <body>
                <p class="text-xs text-left p-3 print:hidden">made by mahmoud eltahawy</p>
                {cards}
            </body>
        </html>
    }
    .to_html()
}

#[component]
pub fn Cards(title: String, cards: Vec<Vec<Kv>>) -> impl IntoView {
    let cards = cards
        .into_iter()
        .map(|kvs| {
            let kvs = kvs
                .into_iter()
                .map(|Kv { key, value }| {
                    view! {
                         <div class="flex">
                            <dt class="text-sm px-2 border-l-2 border-dotted font-bold">{key}</dt>
                            <dd class="grow text-sm">{value}</dd>
                        </div>
                    }
                })
                .collect_view();
            view! {
                <div
                    class="break-inside-avoid border-sky-500 border-5 rounded-xl p-2 m-2 text-xl text-center"
                >
                    <h2 class="font-bold font-xl underline">{title.clone()}</h2>
                    <dl class="divide-y divide-white/10">
                        {kvs}
                    </dl>
                </div>

            }
        })
        .collect_view();

    view! {
        <div class="grid grid-cols-3 gap-1">
            {cards}
        </div>
    }
}

#[derive(Clone)]
pub struct Kv {
    pub key: String,
    pub value: String,
}

fn get_cards(
    title_row_index: usize,
    path: &PathBuf,
    sheet: &str,
    columns_indexs: Vec<usize>,
) -> Result<Vec<Vec<Kv>>, calamine::Error> {
    use calamine::{Data, DeError, RangeDeserializerBuilder, Reader, Xlsx, open_workbook};

    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook.worksheet_range(sheet)?;

    let mut iter = RangeDeserializerBuilder::new()
        .has_headers(false)
        .from_range(&range)?;

    let headers = iter
        .nth(Into::<usize>::into(title_row_index) - 1)
        .unwrap_or(Err(DeError::HeaderNotFound(format!(
            "Error number {title_row_index} should contain headers"
        ))))?;

    let mut cards = Vec::new();
    for row in iter {
        let mut kvs = Vec::new();
        let row: Vec<Data> = row?;
        for index in columns_indexs.iter() {
            let header = headers[*index].to_string();
            let value = row[*index].to_string();
            if !header.is_empty() && !value.is_empty() {
                kvs.push(Kv { key: header, value });
            }
        }
        cards.push(kvs);
    }

    Ok(cards)
}
