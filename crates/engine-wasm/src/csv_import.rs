//! Delimited text (CSV, TSV, semicolon- or pipe-separated) as a new .xlsx.
//!
//! A text file opens by becoming a workbook first: one sheet named after the
//! file, each field read the way Excel reads a typed cell (numbers, dates,
//! percentages, booleans, formulas), then exported. From there it takes the
//! same open, edit and patch-on-save path as any .xlsx.
//!
//! Where Excel loses data on a CSV open, this keeps it: digits with leading
//! zeros (ZIP codes, IDs) and digit strings longer than 15 (card and account
//! numbers) stay text.

use ironcalc_base::expressions::parser::Node;
use ironcalc_base::Model;

/// Excel's worksheet limits.
const MAX_ROWS: usize = 1_048_576;
const MAX_COLUMNS: usize = 16_384;
/// Records the delimiter is guessed from.
const SNIFF_RECORDS: usize = 50;
/// Column width in characters: Excel's default, and the most a column grows to fit its text.
const DEFAULT_CHARS: usize = 10;
const MAX_CHARS: usize = 60;
/// Pixels per character of column width (IronCalc's COLUMN_WIDTH_FACTOR).
const PX_PER_CHAR: f64 = 9.0;

/// The text as one sheet of an .xlsx. `delimiter` overrides the guess
/// (a `sep=` first line, as Excel writes, also sets it).
pub fn csv_to_xlsx(bytes: &[u8], sheet_name: &str, delimiter: Option<char>) -> Result<Vec<u8>, String> {
    let mut model = csv_model(bytes, sheet_name, delimiter)?;
    excel_default_font(&mut model);
    let out = ironcalc::export::save_xlsx_to_writer(&model, std::io::Cursor::new(Vec::new())).map_err(|e| format!("{e:?}"))?;
    Ok(out.into_inner())
}

/// Excel's body font, not IronCalc's (Inter 12): a new workbook should open
/// in Excel looking like one made there.
pub fn excel_default_font(model: &mut Model) {
    if let Some(font) = model.workbook.styles.fonts.first_mut() {
        font.name = "Calibri".to_string();
        font.sz = 11;
    }
}

fn csv_model(bytes: &[u8], sheet_name: &str, delimiter: Option<char>) -> Result<Model<'static>, String> {
    let text = decode(bytes);
    let (body, declared) = strip_sep_line(&text);
    let delim = delimiter.or(declared).unwrap_or_else(|| sniff_delimiter(body));
    let delim = u8::try_from(delim as u32).map_err(|_| format!("the delimiter {delim:?} is not a single-byte character"))?;

    let mut model = Model::new_empty("workbook", "en", "UTC", "en")?;
    let name = sheet_title(sheet_name);
    if let Some(current) = model.workbook.get_worksheet_names().first().cloned() {
        if current != name {
            model.rename_sheet(&current, &name)?;
        }
    }

    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delim)
        .has_headers(false)
        .flexible(true)
        .from_reader(body.as_bytes());
    let comma_decimals = delim == b';';
    let mut widths: Vec<usize> = Vec::new();
    let mut has_formulas = false;
    for (r, record) in reader.records().enumerate() {
        let record = record.map_err(|e| format!("line {}: {e}", r + 1))?;
        if r >= MAX_ROWS {
            return Err(format!("it has more than {MAX_ROWS} rows, the most a worksheet holds"));
        }
        if record.len() > MAX_COLUMNS {
            return Err(format!("line {} has {} fields; a worksheet holds {MAX_COLUMNS} columns", r + 1, record.len()));
        }
        if widths.len() < record.len() {
            widths.resize(record.len(), 0);
        }
        for (c, field) in record.iter().enumerate() {
            if field.is_empty() {
                continue;
            }
            widths[c] = widths[c].max(field.chars().count());
            let input = cell_input(field, comma_decimals);
            let formula = input.starts_with('=');
            let (row, column) = (r as i32 + 1, c as i32 + 1);
            let written = model.set_user_input(0, row, column, input);
            if written.is_err() || (formula && !parses(&model, row, column, field)) {
                // A formula that doesn't parse ("=SUM(", "== heading ==") is
                // the text it is, not a broken formula in the file.
                model.set_user_input(0, row, column, format!("'{field}"))?;
            } else {
                has_formulas |= formula;
            }
        }
    }
    for (c, chars) in widths.iter().enumerate() {
        if *chars > DEFAULT_CHARS {
            model.set_column_width(0, c as i32 + 1, ((*chars).min(MAX_CHARS) + 1) as f64 * PX_PER_CHAR)?;
        }
    }
    // Saved values for the formulas, as a file from Excel carries.
    if has_formulas {
        model.evaluate();
    }
    Ok(model)
}

/// Whether the formula just written to this cell of the first sheet parsed
/// as written. The engine closes an open parenthesis for someone typing
/// ("=SUM(" becomes SUM()); a file's text gets no such repair.
fn parses(model: &Model, row: i32, column: i32, formula: &str) -> bool {
    let index = model.workbook.worksheets[0].cell(row, column).and_then(|cell| cell.get_formula());
    let parsed = index.and_then(|i| model.parsed_formulas[0].get(i as usize));
    let mut depth = 0i32;
    let mut in_string = false;
    for ch in formula.chars() {
        match ch {
            '"' => in_string = !in_string,
            '(' if !in_string => depth += 1,
            ')' if !in_string => depth -= 1,
            _ => {}
        }
    }
    depth == 0 && !matches!(parsed, Some((Node::ParseErrorKind { .. }, _)))
}

/// What to type into the cell for one field.
fn cell_input(field: &str, comma_decimals: bool) -> String {
    // A leading apostrophe is part of the data, not Excel's text marker.
    if field.starts_with('\'') {
        return format!("'{field}");
    }
    if field.len() > 1 && field.starts_with('=') {
        return field.to_string();
    }
    let t = field.trim();
    let digits = t.strip_prefix(['-', '+']).unwrap_or(t);
    if !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit()) {
        // 007, 02134: an ID, not a number. Past 15 digits a number loses its tail.
        if (digits.len() > 1 && digits.starts_with('0')) || digits.len() > 15 {
            return format!("'{field}");
        }
        return t.to_string();
    }
    if comma_decimals {
        if let Some(n) = comma_decimal(t) {
            return n;
        }
    }
    // Spaces around a number don't make it text; around words they're kept.
    if t.len() != field.len() && t.starts_with(|c: char| c.is_ascii_digit() || "+-.$€(".contains(c)) {
        return t.to_string();
    }
    field.to_string()
}

/// "1.234,5" → "1234.5" for files from comma-decimal locales (the ones that
/// separate fields with semicolons); None when it isn't such a number.
fn comma_decimal(t: &str) -> Option<String> {
    let (sign, rest) = match t.strip_prefix('-') {
        Some(r) => ("-", r),
        None => ("", t),
    };
    let (int, frac) = rest.split_once(',')?;
    if frac.is_empty() || !frac.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let groups: Vec<&str> = int.split('.').collect();
    let grouped_ok = groups[0].len() <= 3 && groups[1..].iter().all(|g| g.len() == 3);
    if groups.iter().any(|g| g.is_empty() || !g.bytes().all(|b| b.is_ascii_digit())) || (groups.len() > 1 && !grouped_ok) {
        return None;
    }
    Some(format!("{sign}{}.{frac}", groups.concat()))
}

/// Excel's `sep=;` first line: the delimiter, and the text after that line.
fn strip_sep_line(text: &str) -> (&str, Option<char>) {
    let first_end = text.find('\n').map_or(text.len(), |i| i + 1);
    let first = text[..first_end].trim_end_matches(['\r', '\n']);
    if let Some(rest) = first.strip_prefix("sep=") {
        let mut chars = rest.chars();
        if let (Some(d), None) = (chars.next(), chars.next()) {
            return (&text[first_end..], Some(d));
        }
    }
    (text, None)
}

/// The delimiter most of the first records agree on: the candidate whose
/// per-record count repeats on the most records (ties: more fields, then
/// comma, tab, semicolon, pipe). A file with none of them is one column.
fn sniff_delimiter(text: &str) -> char {
    const CANDIDATES: [char; 4] = [',', '\t', ';', '|'];
    let mut counts: Vec<[usize; 4]> = Vec::new();
    let mut current = [0usize; 4];
    let mut in_quotes = false;
    let mut any = false;
    for ch in text.chars() {
        match ch {
            '"' => in_quotes = !in_quotes,
            '\n' if !in_quotes => {
                if any {
                    counts.push(current);
                    if counts.len() == SNIFF_RECORDS {
                        break;
                    }
                }
                current = [0; 4];
                any = false;
                continue;
            }
            _ if !in_quotes => {
                if let Some(i) = CANDIDATES.iter().position(|&d| d == ch) {
                    current[i] += 1;
                }
            }
            _ => {}
        }
        any |= ch != '\r';
    }
    if any && counts.len() < SNIFF_RECORDS {
        counts.push(current);
    }
    let mut best: Option<(usize, usize, char)> = None; // (records agreeing, count, delimiter)
    for (i, &d) in CANDIDATES.iter().enumerate() {
        let mut tally: Vec<(usize, usize)> = Vec::new(); // (count, records)
        for rec in &counts {
            let n = rec[i];
            if n == 0 {
                continue;
            }
            match tally.iter_mut().find(|(c, _)| *c == n) {
                Some(t) => t.1 += 1,
                None => tally.push((n, 1)),
            }
        }
        let Some(&(count, agree)) = tally.iter().max_by_key(|(c, recs)| (*recs, *c)) else { continue };
        if best.map_or(true, |(a, n, _)| (agree, count) > (a, n)) {
            best = Some((agree, count, d));
        }
    }
    best.map_or(',', |(_, _, d)| d)
}

/// The bytes as text: UTF-8 (with or without a byte-order mark), UTF-16
/// with a mark (Excel's "Unicode Text"), else Windows-1252 (Excel's CSV on
/// Windows in Western locales).
fn decode(bytes: &[u8]) -> String {
    if let Some(rest) = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]) {
        return String::from_utf8_lossy(rest).into_owned();
    }
    if let Some(rest) = bytes.strip_prefix(&[0xFF, 0xFE]) {
        return utf16(rest, u16::from_le_bytes);
    }
    if let Some(rest) = bytes.strip_prefix(&[0xFE, 0xFF]) {
        return utf16(rest, u16::from_be_bytes);
    }
    match std::str::from_utf8(bytes) {
        Ok(s) => s.to_string(),
        Err(_) => bytes.iter().map(|&b| windows_1252(b)).collect(),
    }
}

fn utf16(bytes: &[u8], unit: fn([u8; 2]) -> u16) -> String {
    let units = bytes.chunks_exact(2).map(|p| unit([p[0], p[1]]));
    char::decode_utf16(units).map(|c| c.unwrap_or('\u{FFFD}')).collect()
}

fn windows_1252(b: u8) -> char {
    const HIGH: [u16; 32] = [
        0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F,
        0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
    ];
    match b {
        0x80..=0x9F => char::from_u32(HIGH[(b - 0x80) as usize] as u32).unwrap_or('\u{FFFD}'),
        _ => b as char,
    }
}

/// A valid sheet name from a file name: no extension, none of `[]:*?/\`,
/// at most 31 characters, as Excel names the sheet of a CSV it opens.
fn sheet_title(file_name: &str) -> String {
    let base = file_name.rsplit(['/', '\\']).next().unwrap_or(file_name);
    let stem = match base.rsplit_once('.') {
        Some((s, _)) if !s.is_empty() => s,
        _ => base,
    };
    let clean: String = stem.chars().filter(|c| !"[]:*?/\\".contains(*c)).take(31).collect();
    let clean = clean.trim().trim_matches('\'').trim();
    if clean.is_empty() || clean.eq_ignore_ascii_case("history") {
        "Sheet1".to_string()
    } else {
        clean.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use ironcalc_base::cell::CellValue;

    fn model(text: &str) -> Model<'static> {
        csv_model(text.as_bytes(), "data.csv", None).unwrap()
    }
    fn value(m: &Model, row: i32, col: i32) -> CellValue {
        m.get_cell_value_by_index(0, row, col).unwrap()
    }
    fn text(s: &str) -> CellValue {
        CellValue::String(s.to_string())
    }

    #[test]
    fn values_formulas_and_quotes() {
        let m = model("name,qty,price,total\r\n\"Widget, large\",3,$4.50,=B2*C2\r\nGadget,10,12%,TRUE\r\n");
        assert_eq!(m.workbook.get_worksheet_names(), vec!["data"]);
        assert_eq!(value(&m, 2, 1), text("Widget, large"));
        assert_eq!(value(&m, 2, 2), CellValue::Number(3.0));
        assert_eq!(value(&m, 2, 3), CellValue::Number(4.5));
        assert_eq!(value(&m, 2, 4), CellValue::Number(13.5));
        assert_eq!(value(&m, 3, 3), CellValue::Number(0.12));
        assert_eq!(value(&m, 3, 4), CellValue::Boolean(true));
    }

    #[test]
    fn keeps_what_excel_drops() {
        let m = model("zip,card,quote,neg\n02134,4111111111111111,'hello,-007\n");
        assert_eq!(value(&m, 2, 1), text("02134"));
        assert_eq!(value(&m, 2, 2), text("4111111111111111"));
        assert_eq!(value(&m, 2, 3), text("'hello"));
        assert_eq!(value(&m, 2, 4), text("-007"));
    }

    #[test]
    fn delimiters() {
        assert_eq!(sniff_delimiter("a\tb\tc\n1\t2\t3\n"), '\t');
        assert_eq!(sniff_delimiter("a;b;c\n1,5;2,25;3\n"), ';');
        assert_eq!(sniff_delimiter("a|b\n1|2\n"), '|');
        assert_eq!(sniff_delimiter("\"x,y\";b\n\"1,2\";3\n"), ';');
        assert_eq!(sniff_delimiter("just one column\nsecond\n"), ',');
        let m = model("sep=;\nprice;qty\n1.234,5;-0,75\n");
        assert_eq!(value(&m, 1, 2), text("qty"));
        assert_eq!(value(&m, 2, 1), CellValue::Number(1234.5));
        assert_eq!(value(&m, 2, 2), CellValue::Number(-0.75));
    }

    #[test]
    fn encodings() {
        let mut utf16 = vec![0xFF, 0xFE];
        for u in "é\t€\n1\t2\n".encode_utf16() {
            utf16.extend(u.to_le_bytes());
        }
        let m = csv_model(&utf16, "x.txt", None).unwrap();
        assert_eq!(value(&m, 1, 2), text("€"));
        let m = csv_model(b"caf\xe9,\x80\n", "x.csv", None).unwrap();
        assert_eq!(value(&m, 1, 1), text("café"));
        assert_eq!(value(&m, 1, 2), text("€"));
        let m = csv_model(b"\xEF\xBB\xBFa,b\n", "x.csv", None).unwrap();
        assert_eq!(value(&m, 1, 1), text("a"));
    }

    #[test]
    fn ragged_rows_bad_formulas_and_names() {
        let m = model("a\n1,2,3\n=SUM(\n== notes ==,=1+\n");
        assert_eq!(value(&m, 2, 3), CellValue::Number(3.0));
        assert_eq!(value(&m, 3, 1), text("=SUM("));
        assert_eq!(value(&m, 4, 1), text("== notes =="));
        assert_eq!(value(&m, 4, 2), text("=1+"));
        assert_eq!(sheet_title("C:\\exports\\Q3 [final].csv"), "Q3 final");
        assert_eq!(sheet_title("a-very-long-file-name-that-goes-past-31.csv").chars().count(), 31);
        assert_eq!(sheet_title(".csv"), ".csv");
        assert_eq!(sheet_title("history.csv"), "Sheet1");
    }

    #[test]
    fn exports_an_xlsx_that_reopens() {
        let bytes = csv_to_xlsx(b"a,b\n1,=A2*2\n", "nums.csv", None).unwrap();
        assert_eq!(&bytes[..2], b"PK");
        let wb = ironcalc::import::load_from_xlsx_bytes(&bytes, "nums", "en", "UTC").unwrap();
        let m = Model::from_workbook(wb, "en").unwrap();
        assert_eq!(m.workbook.get_worksheet_names(), vec!["nums"]);
        assert_eq!(m.get_cell_value_by_index(0, 2, 2).unwrap(), CellValue::Number(2.0));
    }
}
