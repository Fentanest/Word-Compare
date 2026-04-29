use std::collections::HashMap;
use std::env;
use std::fs::File;
use std::io::{Read, Write};

use roxmltree::{Document, Node};
use serde::Serialize;
use zip::ZipArchive;

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

fn main() {
    if let Err(error) = run() {
        let _ = writeln!(std::io::stderr(), "{error}");
        std::process::exit(1);
    }
}

fn run() -> Result<(), String> {
    let input_path = env::args()
        .nth(1)
        .ok_or_else(|| "usage: word_compare_native_extractor <input.docx>".to_string())?;

    let package = DocxPackage::load(&input_path)?;
    let payload = build_payload(&package)?;
    let stdout = std::io::stdout();
    let mut handle = stdout.lock();
    serde_json::to_writer(&mut handle, &payload).map_err(|error| error.to_string())?;
    handle.write_all(b"\n").map_err(|error| error.to_string())?;
    Ok(())
}

#[derive(Debug)]
struct DocxPackage {
    parts: HashMap<String, String>,
    names: Vec<String>,
}

impl DocxPackage {
    fn load(path: &str) -> Result<Self, String> {
        let file = File::open(path).map_err(|error| format!("failed to open {path}: {error}"))?;
        let mut archive = ZipArchive::new(file).map_err(|error| format!("failed to open zip: {error}"))?;
        let mut parts = HashMap::new();
        let mut names = Vec::new();

        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).map_err(|error| error.to_string())?;
            let name = entry.name().to_string();
            names.push(name.clone());
            if !name.ends_with(".xml") {
                continue;
            }

            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes).map_err(|error| error.to_string())?;
            let text = String::from_utf8(bytes).map_err(|error| error.to_string())?;
            parts.insert(name, text);
        }

        names.sort();
        Ok(Self { parts, names })
    }

    fn xml(&self, part_name: &str) -> Result<Option<Document<'_>>, String> {
        match self.parts.get(part_name) {
            Some(xml_text) => Document::parse(xml_text)
                .map(Some)
                .map_err(|error| format!("failed to parse {part_name}: {error}")),
            None => Ok(None),
        }
    }

    fn matching_names(&self, prefix: &str) -> Vec<String> {
        self.names
            .iter()
            .filter(|name| name.starts_with(prefix) && name.ends_with(".xml"))
            .cloned()
            .collect()
    }
}

#[derive(Serialize)]
struct ExtractedDocumentPayload {
    paragraphs: Vec<ParagraphPayload>,
    table_flags: Vec<bool>,
    tables: Vec<Vec<Vec<TableCellPayload>>>,
    paragraph_locations: Vec<String>,
}

#[derive(Serialize)]
struct ParagraphPayload {
    #[serde(rename = "type")]
    type_name: String,
    text: String,
    style_name: String,
    alignment: String,
    source_kind: String,
    source_identifier: String,
    left_indent: i32,
    right_indent: i32,
    first_line_indent: i32,
    space_before: i32,
    space_after: i32,
    line_spacing: String,
    keep_together: bool,
    keep_with_next: bool,
    page_break_before: bool,
    widow_control: bool,
    runs: Vec<RunPayload>,
    extra_meta: Vec<String>,
}

#[derive(Serialize)]
struct RunPayload {
    text: String,
    bold: bool,
    italic: bool,
    underline: String,
    font_name: String,
    font_size: i32,
    color: String,
    highlight: String,
    strike: bool,
    style_name: String,
}

#[derive(Serialize)]
struct TableCellPayload {
    text: String,
    grid_span: i32,
    v_merge: String,
    cell_width: i32,
    row_height: i32,
    grid_col_width: i32,
    border_signature: String,
    shading_fill: String,
    vertical_align: String,
    text_direction: String,
    nested_table_count: i32,
}

fn build_payload(package: &DocxPackage) -> Result<ExtractedDocumentPayload, String> {
    let document = package
        .xml("word/document.xml")?
        .ok_or_else(|| "word/document.xml is missing".to_string())?;
    let body = document
        .descendants()
        .find(|node| is_w_tag(*node, "body"))
        .ok_or_else(|| "word/document.xml does not contain w:body".to_string())?;

    let mut paragraphs = Vec::new();
    let mut table_flags = Vec::new();
    let mut tables = Vec::new();
    let mut paragraph_locations = Vec::new();

    for child in body.children().filter(|node| node.is_element()) {
        match local_name(child) {
            Some("p") => {
                paragraphs.push(build_body_paragraph(child));
                table_flags.push(false);
                paragraph_locations.push(format!("{}행", paragraph_locations.len() + 1));
            }
            Some("tbl") => {
                paragraphs.push(marker_paragraph("[TABLE_MARKER]"));
                table_flags.push(true);
                paragraph_locations.push(format!("{}행", paragraph_locations.len() + 1));
                tables.push(build_table(child));
            }
            _ => {}
        }
    }

    let (metadata_blocks, metadata_locations) = extract_metadata_blocks(package)?;
    paragraphs.extend(metadata_blocks);
    table_flags.extend(std::iter::repeat(false).take(metadata_locations.len()));
    paragraph_locations.extend(metadata_locations);

    Ok(ExtractedDocumentPayload {
        paragraphs,
        table_flags,
        tables,
        paragraph_locations,
    })
}

fn build_body_paragraph(node: Node<'_, '_>) -> ParagraphPayload {
    let paragraph_pr = child_w(node, "pPr");
    let spacing = paragraph_pr.and_then(|pr| child_w(pr, "spacing"));
    let indent = paragraph_pr.and_then(|pr| child_w(pr, "ind"));

    ParagraphPayload {
        type_name: "paragraph".to_string(),
        text: extract_text_from_element(node),
        style_name: paragraph_pr
            .and_then(|pr| child_w(pr, "pStyle"))
            .and_then(|node| w_attr(node, "val"))
            .unwrap_or_default(),
        alignment: paragraph_pr
            .and_then(|pr| child_w(pr, "jc"))
            .and_then(|node| w_attr(node, "val"))
            .unwrap_or_default(),
        source_kind: "body".to_string(),
        source_identifier: String::new(),
        left_indent: indent
            .and_then(|node| parse_i32_opt(w_attr(node, "left")))
            .unwrap_or(0),
        right_indent: indent
            .and_then(|node| parse_i32_opt(w_attr(node, "right")))
            .unwrap_or(0),
        first_line_indent: parse_first_line_indent(indent),
        space_before: spacing
            .and_then(|node| parse_i32_opt(w_attr(node, "before")))
            .unwrap_or(0),
        space_after: spacing
            .and_then(|node| parse_i32_opt(w_attr(node, "after")))
            .unwrap_or(0),
        line_spacing: spacing
            .and_then(|node| w_attr(node, "line"))
            .unwrap_or_default(),
        keep_together: paragraph_pr.map_or(false, |pr| child_w(pr, "keepLines").is_some()),
        keep_with_next: paragraph_pr.map_or(false, |pr| child_w(pr, "keepNext").is_some()),
        page_break_before: paragraph_pr.map_or(false, |pr| child_w(pr, "pageBreakBefore").is_some()),
        widow_control: paragraph_pr.map_or(false, |pr| child_w(pr, "widowControl").is_some()),
        runs: node
            .children()
            .filter(|child| is_w_tag(*child, "r"))
            .map(build_run)
            .filter(|run| !run.text.is_empty())
            .collect(),
        extra_meta: Vec::new(),
    }
}

fn build_run(node: Node<'_, '_>) -> RunPayload {
    let run_pr = child_w(node, "rPr");
    let fonts = run_pr.and_then(|pr| child_w(pr, "rFonts"));
    RunPayload {
        text: extract_text_from_element(node),
        bold: run_pr.map_or(false, |pr| toggle_true(child_w(pr, "b"))),
        italic: run_pr.map_or(false, |pr| toggle_true(child_w(pr, "i"))),
        underline: run_pr
            .and_then(|pr| child_w(pr, "u"))
            .and_then(|node| w_attr(node, "val"))
            .unwrap_or_default(),
        font_name: fonts
            .and_then(|node| w_attr(node, "ascii").or_else(|| w_attr(node, "hAnsi")))
            .unwrap_or_default(),
        font_size: run_pr
            .and_then(|pr| child_w(pr, "sz"))
            .and_then(|node| parse_i32_opt(w_attr(node, "val")))
            .unwrap_or(0),
        color: run_pr
            .and_then(|pr| child_w(pr, "color"))
            .and_then(|node| w_attr(node, "val"))
            .unwrap_or_default(),
        highlight: run_pr
            .and_then(|pr| child_w(pr, "highlight"))
            .and_then(|node| w_attr(node, "val"))
            .unwrap_or_default(),
        strike: run_pr.map_or(false, |pr| toggle_true(child_w(pr, "strike"))),
        style_name: run_pr
            .and_then(|pr| child_w(pr, "rStyle"))
            .and_then(|node| w_attr(node, "val"))
            .unwrap_or_default(),
    }
}

fn build_table(table_node: Node<'_, '_>) -> Vec<Vec<TableCellPayload>> {
    let grid_widths = extract_table_grid_widths(table_node);
    table_node
        .children()
        .filter(|node| is_w_tag(*node, "tr"))
        .map(|row_node| build_table_row(row_node, &grid_widths))
        .collect()
}

fn build_table_row(row_node: Node<'_, '_>, grid_widths: &[i32]) -> Vec<TableCellPayload> {
    let row_height = child_w(row_node, "trPr")
        .and_then(|node| child_w(node, "trHeight"))
        .and_then(|node| parse_i32_opt(w_attr(node, "val")))
        .unwrap_or(0);

    row_node
        .children()
        .filter(|node| is_w_tag(*node, "tc"))
        .enumerate()
        .map(|(cell_index, cell_node)| {
            let tc_pr = child_w(cell_node, "tcPr");
            let grid_col_width = grid_widths.get(cell_index).copied().unwrap_or(0);

            TableCellPayload {
                text: extract_text_from_element(cell_node),
                grid_span: tc_pr
                    .and_then(|node| child_w(node, "gridSpan"))
                    .and_then(|node| parse_i32_opt(w_attr(node, "val")))
                    .unwrap_or(1),
                v_merge: tc_pr
                    .and_then(|node| child_w(node, "vMerge"))
                    .and_then(|node| w_attr(node, "val"))
                    .unwrap_or_else(|| {
                        if tc_pr.and_then(|node| child_w(node, "vMerge")).is_some() {
                            "continue".to_string()
                        } else {
                            String::new()
                        }
                    }),
                cell_width: tc_pr
                    .and_then(|node| child_w(node, "tcW"))
                    .and_then(|node| parse_i32_opt(w_attr(node, "w")))
                    .unwrap_or(0),
                row_height,
                grid_col_width,
                border_signature: tc_pr
                    .and_then(|node| child_w(node, "tcBorders"))
                    .map(border_signature)
                    .unwrap_or_default(),
                shading_fill: tc_pr
                    .and_then(|node| child_w(node, "shd"))
                    .map(shading_signature)
                    .unwrap_or_default(),
                vertical_align: tc_pr
                    .and_then(|node| child_w(node, "vAlign"))
                    .and_then(|node| w_attr(node, "val"))
                    .unwrap_or_default(),
                text_direction: tc_pr
                    .and_then(|node| child_w(node, "textDirection"))
                    .and_then(|node| w_attr(node, "val"))
                    .unwrap_or_default(),
                nested_table_count: cell_node
                    .descendants()
                    .filter(|node| is_w_tag(*node, "tbl"))
                    .count()
                    .saturating_sub(1) as i32,
            }
        })
        .collect()
}

fn extract_metadata_blocks(package: &DocxPackage) -> Result<(Vec<ParagraphPayload>, Vec<String>), String> {
    let mut blocks = Vec::new();
    let mut locations = Vec::new();

    append_header_footer_blocks(package, &mut blocks, &mut locations, "word/header", "header", "머리말")?;
    append_header_footer_blocks(package, &mut blocks, &mut locations, "word/footer", "footer", "꼬리말")?;
    append_note_blocks(package, &mut blocks, &mut locations, "word/footnotes.xml", "footnote", "각주")?;
    append_note_blocks(package, &mut blocks, &mut locations, "word/endnotes.xml", "endnote", "미주")?;
    append_comment_blocks(package, &mut blocks, &mut locations)?;
    append_revision_blocks(package, &mut blocks, &mut locations)?;
    append_section_blocks(package, &mut blocks, &mut locations)?;
    append_shape_blocks(package, &mut blocks, &mut locations)?;
    append_table_metadata_blocks(package, &mut blocks, &mut locations)?;

    Ok((blocks, locations))
}

fn append_header_footer_blocks(
    package: &DocxPackage,
    blocks: &mut Vec<ParagraphPayload>,
    locations: &mut Vec<String>,
    prefix: &str,
    kind: &str,
    label_prefix: &str,
) -> Result<(), String> {
    for (index, part_name) in package.matching_names(prefix).into_iter().enumerate() {
        if let Some(document) = package.xml(&part_name)? {
            let text = extract_paragraph_texts(document.root_element()).join("\n").trim().to_string();
            push_metadata_block(
                blocks,
                locations,
                format!("{label_prefix} {}", index + 1),
                if text.is_empty() {
                    format!("{label_prefix} {}", index + 1)
                } else {
                    text
                },
                kind.to_string(),
                part_name,
                Vec::new(),
            );
        }
    }
    Ok(())
}

fn append_note_blocks(
    package: &DocxPackage,
    blocks: &mut Vec<ParagraphPayload>,
    locations: &mut Vec<String>,
    part_name: &str,
    tag_name: &str,
    label_prefix: &str,
) -> Result<(), String> {
    let Some(document) = package.xml(part_name)? else {
        return Ok(());
    };

    for note in document.descendants().filter(|node| is_w_tag(*node, tag_name)) {
        let note_type = w_attr(note, "type").unwrap_or_default();
        if !note_type.is_empty() {
            continue;
        }

        let note_id = w_attr(note, "id").unwrap_or_else(|| "0".to_string());
        if note_id.parse::<i32>().ok().is_some_and(|value| value < 0) {
            continue;
        }

        let text = extract_paragraph_texts(note).join("\n").trim().to_string();
        push_metadata_block(
            blocks,
            locations,
            format!("{label_prefix} {note_id}"),
            if text.is_empty() {
                format!("{label_prefix} {note_id}")
            } else {
                text
            },
            tag_name.to_string(),
            note_id.clone(),
            vec![format!("part={part_name}")],
        );
    }
    Ok(())
}

fn append_comment_blocks(
    package: &DocxPackage,
    blocks: &mut Vec<ParagraphPayload>,
    locations: &mut Vec<String>,
) -> Result<(), String> {
    let Some(document) = package.xml("word/comments.xml")? else {
        return Ok(());
    };

    for comment in document.descendants().filter(|node| is_w_tag(*node, "comment")) {
        let comment_id = w_attr(comment, "id").unwrap_or_else(|| "0".to_string());
        let text = extract_paragraph_texts(comment).join("\n").trim().to_string();
        push_metadata_block(
            blocks,
            locations,
            format!("주석 {comment_id}"),
            if text.is_empty() {
                format!("주석 {comment_id}")
            } else {
                text
            },
            "comment".to_string(),
            comment_id,
            vec![
                format!("author={}", w_attr(comment, "author").unwrap_or_default()),
                format!("initials={}", w_attr(comment, "initials").unwrap_or_default()),
                format!("date={}", w_attr(comment, "date").unwrap_or_default()),
            ],
        );
    }
    Ok(())
}

fn append_revision_blocks(
    package: &DocxPackage,
    blocks: &mut Vec<ParagraphPayload>,
    locations: &mut Vec<String>,
) -> Result<(), String> {
    let tag_labels = [
        ("ins", "삽입"),
        ("del", "삭제"),
        ("moveFrom", "이동-출발"),
        ("moveTo", "이동-도착"),
    ];
    let mut counters: HashMap<&str, usize> = tag_labels.iter().map(|(name, _)| (*name, 0)).collect();
    let mut part_names = vec!["word/document.xml".to_string()];
    part_names.extend(package.matching_names("word/header"));
    part_names.extend(package.matching_names("word/footer"));
    if package.parts.contains_key("word/footnotes.xml") {
        part_names.push("word/footnotes.xml".to_string());
    }
    if package.parts.contains_key("word/endnotes.xml") {
        part_names.push("word/endnotes.xml".to_string());
    }
    part_names.sort();
    part_names.dedup();

    for part_name in part_names {
        let Some(document) = package.xml(&part_name)? else {
            continue;
        };
        for (tag_name, label) in tag_labels {
            for revision in document.descendants().filter(|node| is_w_tag(*node, tag_name)) {
                let counter = counters.entry(tag_name).or_insert(0);
                *counter += 1;
                let text = extract_text_from_element(revision);
                push_metadata_block(
                    blocks,
                    locations,
                    format!("변경 추적 {label} {counter}"),
                    if text.is_empty() {
                        format!("변경 추적 {label}")
                    } else {
                        text
                    },
                    "revision".to_string(),
                    format!("{tag_name}:{counter}"),
                    vec![
                        format!("part={part_name}"),
                        format!("id={}", w_attr(revision, "id").unwrap_or_default()),
                        format!("author={}", w_attr(revision, "author").unwrap_or_default()),
                        format!("date={}", w_attr(revision, "date").unwrap_or_default()),
                        format!("type={tag_name}"),
                    ],
                );
            }
        }
    }
    Ok(())
}

fn append_section_blocks(
    package: &DocxPackage,
    blocks: &mut Vec<ParagraphPayload>,
    locations: &mut Vec<String>,
) -> Result<(), String> {
    let Some(document) = package.xml("word/document.xml")? else {
        return Ok(());
    };

    for (index, sect_pr) in document
        .descendants()
        .filter(|node| is_w_tag(*node, "sectPr"))
        .enumerate()
    {
        push_metadata_block(
            blocks,
            locations,
            format!("구역 {}", index + 1),
            format!("구역 {} 설정", index + 1),
            "section".to_string(),
            (index + 1).to_string(),
            extract_section_tokens(sect_pr),
        );
    }
    Ok(())
}

fn append_shape_blocks(
    package: &DocxPackage,
    blocks: &mut Vec<ParagraphPayload>,
    locations: &mut Vec<String>,
) -> Result<(), String> {
    let mut part_names = vec!["word/document.xml".to_string()];
    part_names.extend(package.matching_names("word/header"));
    part_names.extend(package.matching_names("word/footer"));
    part_names.sort();
    part_names.dedup();

    let mut textbox_counter = 0usize;
    let mut shape_counter = 0usize;

    for part_name in part_names {
        let Some(document) = package.xml(&part_name)? else {
            continue;
        };

        for textbox in document.descendants().filter(|node| is_w_tag(*node, "txbxContent")) {
            textbox_counter += 1;
            let text = extract_paragraph_texts(textbox).join("\n").trim().to_string();
            push_metadata_block(
                blocks,
                locations,
                format!("텍스트 상자 {textbox_counter}"),
                if text.is_empty() {
                    format!("텍스트 상자 {textbox_counter}")
                } else {
                    text
                },
                "textbox".to_string(),
                format!("{part_name}:{textbox_counter}"),
                vec![format!("part={part_name}")],
            );
        }

        for doc_pr in document.descendants().filter(|node| is_tag(*node, WP_NS, "docPr")) {
            shape_counter += 1;
            let name = doc_pr.attribute("name").unwrap_or_default();
            let descr = doc_pr.attribute("descr").unwrap_or_default();
            let title = doc_pr.attribute("title").unwrap_or_default();
            let visible_text = first_non_empty(&[name, descr, title]).unwrap_or_else(|| format!("도형 {shape_counter}"));
            push_metadata_block(
                blocks,
                locations,
                format!("도형 {shape_counter}"),
                visible_text,
                "shape".to_string(),
                format!("{part_name}:{shape_counter}"),
                vec![
                    format!("part={part_name}"),
                    format!("name={name}"),
                    format!("descr={descr}"),
                    format!("title={title}"),
                ],
            );
        }
    }
    Ok(())
}

fn append_table_metadata_blocks(
    package: &DocxPackage,
    blocks: &mut Vec<ParagraphPayload>,
    locations: &mut Vec<String>,
) -> Result<(), String> {
    let Some(document) = package.xml("word/document.xml")? else {
        return Ok(());
    };

    for (index, table) in document.descendants().filter(|node| is_w_tag(*node, "tbl")).enumerate() {
        push_metadata_block(
            blocks,
            locations,
            format!("표 {} 서식", index + 1),
            format!("표 {} 서식", index + 1),
            "table-meta".to_string(),
            (index + 1).to_string(),
            extract_table_tokens(table),
        );
    }
    Ok(())
}

fn push_metadata_block(
    blocks: &mut Vec<ParagraphPayload>,
    locations: &mut Vec<String>,
    label: String,
    text: String,
    kind: String,
    identifier: String,
    extra_meta: Vec<String>,
) {
    blocks.push(ParagraphPayload {
        type_name: "paragraph".to_string(),
        text,
        style_name: String::new(),
        alignment: String::new(),
        source_kind: kind,
        source_identifier: identifier,
        left_indent: 0,
        right_indent: 0,
        first_line_indent: 0,
        space_before: 0,
        space_after: 0,
        line_spacing: String::new(),
        keep_together: false,
        keep_with_next: false,
        page_break_before: false,
        widow_control: false,
        runs: Vec::new(),
        extra_meta: extra_meta.into_iter().filter(|item| !item.is_empty() && item != "=").collect(),
    });
    locations.push(label);
}

fn marker_paragraph(text: &str) -> ParagraphPayload {
    ParagraphPayload {
        type_name: "marker".to_string(),
        text: text.to_string(),
        style_name: String::new(),
        alignment: String::new(),
        source_kind: String::new(),
        source_identifier: String::new(),
        left_indent: 0,
        right_indent: 0,
        first_line_indent: 0,
        space_before: 0,
        space_after: 0,
        line_spacing: String::new(),
        keep_together: false,
        keep_with_next: false,
        page_break_before: false,
        widow_control: false,
        runs: Vec::new(),
        extra_meta: Vec::new(),
    }
}

fn extract_table_grid_widths(table_node: Node<'_, '_>) -> Vec<i32> {
    child_w(table_node, "tblGrid")
        .map(|grid| {
            grid.children()
                .filter(|node| is_w_tag(*node, "gridCol"))
                .map(|node| parse_i32_opt(w_attr(node, "w")).unwrap_or(0))
                .collect()
        })
        .unwrap_or_default()
}

fn extract_paragraph_texts(node: Node<'_, '_>) -> Vec<String> {
    let mut texts: Vec<String> = node
        .descendants()
        .filter(|child| is_w_tag(*child, "p"))
        .map(extract_text_from_element)
        .filter(|text| !text.is_empty())
        .collect();

    if texts.is_empty() {
        let fallback = extract_text_from_element(node);
        if !fallback.is_empty() {
            texts.push(fallback);
        }
    }
    texts
}

fn extract_text_from_element(node: Node<'_, '_>) -> String {
    let mut parts = Vec::new();
    for child in node.descendants().filter(|child| child.is_element()) {
        match local_name(child) {
            Some("t") | Some("delText") => {
                if let Some(text) = child.text() {
                    parts.push(text.to_string());
                }
            }
            Some("tab") => parts.push("\t".to_string()),
            Some("br") | Some("cr") => parts.push("\n".to_string()),
            _ => {}
        }
    }
    parts.join("").trim().to_string()
}

fn extract_section_tokens(sect_pr: Node<'_, '_>) -> Vec<String> {
    let mut tokens = Vec::new();
    if let Some(pg_sz) = child_w(sect_pr, "pgSz") {
        push_attr_token(&mut tokens, "pgSz:w", w_attr(pg_sz, "w"));
        push_attr_token(&mut tokens, "pgSz:h", w_attr(pg_sz, "h"));
        push_attr_token(&mut tokens, "pgSz:orient", w_attr(pg_sz, "orient"));
    }

    if let Some(pg_mar) = child_w(sect_pr, "pgMar") {
        for edge in ["top", "right", "bottom", "left", "header", "footer", "gutter"] {
            push_attr_token(&mut tokens, &format!("pgMar:{edge}"), w_attr(pg_mar, edge));
        }
    }

    if let Some(cols) = child_w(sect_pr, "cols") {
        for attribute in ["num", "space", "sep", "equalWidth"] {
            push_attr_token(&mut tokens, &format!("cols:{attribute}"), w_attr(cols, attribute));
        }
    }

    if let Some(pg_num_type) = child_w(sect_pr, "pgNumType") {
        for attribute in ["start", "fmt", "chapStyle"] {
            push_attr_token(&mut tokens, &format!("pgNumType:{attribute}"), w_attr(pg_num_type, attribute));
        }
    }

    if let Some(doc_grid) = child_w(sect_pr, "docGrid") {
        for attribute in ["type", "linePitch", "charSpace"] {
            push_attr_token(&mut tokens, &format!("docGrid:{attribute}"), w_attr(doc_grid, attribute));
        }
    }

    if child_w(sect_pr, "titlePg").is_some() {
        tokens.push("titlePg=true".to_string());
    }
    tokens
}

fn extract_table_tokens(table_node: Node<'_, '_>) -> Vec<String> {
    let mut tokens = Vec::new();
    let Some(table_pr) = child_w(table_node, "tblPr") else {
        return tokens;
    };

    for tag_name in ["tblStyle", "tblW", "jc", "tblLayout"] {
        if let Some(element) = child_w(table_pr, tag_name) {
            for attribute in ["val", "w", "type"] {
                push_attr_token(
                    &mut tokens,
                    &format!("{tag_name}:{attribute}"),
                    w_attr(element, attribute),
                );
            }
        }
    }

    if let Some(tbl_look) = child_w(table_pr, "tblLook") {
        for attribute in ["val", "firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand"] {
            push_attr_token(
                &mut tokens,
                &format!("tblLook:{attribute}"),
                w_attr(tbl_look, attribute),
            );
        }
    }

    if let Some(tbl_borders) = child_w(table_pr, "tblBorders") {
        tokens.extend(extract_border_tokens(tbl_borders, "tblBorders"));
    }
    if let Some(tbl_cell_mar) = child_w(table_pr, "tblCellMar") {
        tokens.extend(extract_margin_tokens(tbl_cell_mar, "tblCellMar"));
    }
    if let Some(shd) = child_w(table_pr, "shd") {
        for attribute in ["val", "color", "fill"] {
            push_attr_token(&mut tokens, &format!("tblShd:{attribute}"), w_attr(shd, attribute));
        }
    }

    tokens
}

fn extract_border_tokens(borders: Node<'_, '_>, prefix: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    for edge_name in ["top", "left", "bottom", "right", "insideH", "insideV"] {
        if let Some(edge) = child_w(borders, edge_name) {
            for attribute in ["val", "sz", "space", "color"] {
                push_attr_token(
                    &mut tokens,
                    &format!("{prefix}:{edge_name}:{attribute}"),
                    w_attr(edge, attribute),
                );
            }
        }
    }
    tokens
}

fn extract_margin_tokens(margins: Node<'_, '_>, prefix: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    for edge_name in ["top", "left", "bottom", "right"] {
        if let Some(edge) = child_w(margins, edge_name) {
            for attribute in ["w", "type"] {
                push_attr_token(
                    &mut tokens,
                    &format!("{prefix}:{edge_name}:{attribute}"),
                    w_attr(edge, attribute),
                );
            }
        }
    }
    tokens
}

fn border_signature(borders: Node<'_, '_>) -> String {
    let mut parts = Vec::new();
    for edge_name in ["top", "left", "bottom", "right", "insideH", "insideV", "tl2br", "tr2bl"] {
        if let Some(edge) = child_w(borders, edge_name) {
            let mut edge_bits = Vec::new();
            for attribute in ["val", "sz", "space", "color"] {
                if let Some(value) = w_attr(edge, attribute) {
                    edge_bits.push(format!("{attribute}={value}"));
                }
            }
            if !edge_bits.is_empty() {
                parts.push(format!("{edge_name}:{}", edge_bits.join(",")));
            }
        }
    }
    parts.join("|")
}

fn shading_signature(shading: Node<'_, '_>) -> String {
    ["val", "color", "fill"]
        .iter()
        .filter_map(|attribute| w_attr(shading, attribute).map(|value| format!("{attribute}={value}")))
        .collect::<Vec<_>>()
        .join("|")
}

fn parse_first_line_indent(indent: Option<Node<'_, '_>>) -> i32 {
    let Some(indent) = indent else {
        return 0;
    };
    if let Some(first_line) = w_attr(indent, "firstLine").and_then(parse_i32_value) {
        return first_line;
    }
    if let Some(hanging) = w_attr(indent, "hanging").and_then(parse_i32_value) {
        return -hanging;
    }
    0
}

fn toggle_true(node: Option<Node<'_, '_>>) -> bool {
    let Some(node) = node else {
        return false;
    };
    match w_attr(node, "val") {
        Some(value) => !matches!(value.as_str(), "0" | "false" | "off"),
        None => true,
    }
}

fn push_attr_token(tokens: &mut Vec<String>, prefix: &str, value: Option<String>) {
    if let Some(value) = value {
        if !value.is_empty() {
            tokens.push(format!("{prefix}={value}"));
        }
    }
}

fn child_w<'a, 'input>(node: Node<'a, 'input>, local: &str) -> Option<Node<'a, 'input>> {
    node.children().find(|child| is_w_tag(*child, local))
}

fn w_attr(node: Node<'_, '_>, local: &str) -> Option<String> {
    node.attribute((W_NS, local)).map(|value| value.to_string())
}

fn is_w_tag(node: Node<'_, '_>, local: &str) -> bool {
    is_tag(node, W_NS, local)
}

fn is_tag(node: Node<'_, '_>, namespace: &str, local: &str) -> bool {
    node.is_element()
        && node.tag_name().name() == local
        && node.tag_name().namespace() == Some(namespace)
}

fn local_name<'a, 'input>(node: Node<'a, 'input>) -> Option<&'input str> {
    if node.is_element() {
        Some(node.tag_name().name())
    } else {
        None
    }
}

fn parse_i32_opt(value: Option<String>) -> Option<i32> {
    value.and_then(|value| value.parse::<i32>().ok())
}

fn parse_i32_value(value: String) -> Option<i32> {
    value.parse::<i32>().ok()
}

fn first_non_empty(values: &[&str]) -> Option<String> {
    values
        .iter()
        .find(|value| !value.is_empty())
        .map(|value| (*value).to_string())
}
