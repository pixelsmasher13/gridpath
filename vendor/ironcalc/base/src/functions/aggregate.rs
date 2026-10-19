use std::collections::{HashMap, HashSet};

use crate::{
    calc_result::CalcResult,
    constants::{LAST_COLUMN, LAST_ROW},
    expressions::{
        parser::{ArrayNode, Node},
        token::Error,
        types::CellReferenceIndex,
    },
    functions::Function,
    model::Model,
};

/// What AGGREGATE's `options` argument (0-7) leaves out.
struct AggregateOptions {
    skip_hidden: bool,
    skip_errors: bool,
    skip_nested: bool,
}

impl AggregateOptions {
    fn from_code(code: i32) -> Option<AggregateOptions> {
        let (skip_hidden, skip_errors, skip_nested) = match code {
            0 => (false, false, true),
            1 => (true, false, true),
            2 => (false, true, true),
            3 => (true, true, true),
            4 => (false, false, false),
            5 => (true, false, false),
            6 => (false, true, false),
            7 => (true, true, false),
            _ => return None,
        };
        Some(AggregateOptions {
            skip_hidden,
            skip_errors,
            skip_nested,
        })
    }
}

/// The values AGGREGATE collected: numbers, a count of the other non-empty
/// values (text, booleans, errors) that only COUNTA counts, and the first
/// error kept (options 0, 1, 4, 5), which every function but COUNT and COUNTA
/// returns.
#[derive(Default)]
struct Collected {
    numbers: Vec<f64>,
    others: usize,
    first_error: Option<CalcResult>,
}

impl<'a> Model<'a> {
    /// AGGREGATE(function_num, options, ref1, [ref2], …)   function_num 1-13
    /// AGGREGATE(function_num, options, array, k)          function_num 14-19
    ///
    /// Hidden rows include rows filtered out by an autofilter. Functions 1-13
    /// take references only (an array is #VALUE!, as in Excel; that is why
    /// `AGGREGATE(14,6,rng/(cond),1)` is the usual conditional MAX).
    pub(crate) fn fn_aggregate(&mut self, args: &[Node], cell: CellReferenceIndex) -> CalcResult {
        if args.len() < 3 {
            return CalcResult::new_args_number_error(cell);
        }
        let function_num = match self.get_number(&args[0], cell) {
            Ok(f) => f.trunc() as i32,
            Err(e) => return e,
        };
        let options = match self.get_number(&args[1], cell) {
            Ok(f) => match AggregateOptions::from_code(f.trunc() as i32) {
                Some(o) => o,
                None => {
                    return CalcResult::new_error(
                        Error::VALUE,
                        cell,
                        "Invalid AGGREGATE options".to_string(),
                    )
                }
            },
            Err(e) => return e,
        };
        match function_num {
            1..=13 => {
                let mut collected = Collected::default();
                for arg in &args[2..] {
                    if let Err(e) = self.aggregate_collect(arg, cell, &options, false, &mut collected) {
                        return e;
                    }
                }
                aggregate_reduce(function_num, collected, cell)
            }
            14..=19 => {
                if args.len() != 4 {
                    return CalcResult::new_error(
                        Error::VALUE,
                        cell,
                        "AGGREGATE functions 14-19 take an array and k".to_string(),
                    );
                }
                let mut collected = Collected::default();
                if let Err(e) = self.aggregate_collect(&args[2], cell, &options, true, &mut collected) {
                    return e;
                }
                let k = match self.get_number(&args[3], cell) {
                    Ok(k) => k,
                    Err(e) => return e,
                };
                if let Some(error) = collected.first_error {
                    return error;
                }
                let mut sorted = collected.numbers;
                sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
                aggregate_ranked(function_num, &sorted, k, cell)
            }
            _ => CalcResult::new_error(
                Error::VALUE,
                cell,
                format!("Invalid AGGREGATE function: {function_num}"),
            ),
        }
    }

    fn aggregate_collect(
        &mut self,
        arg: &Node,
        cell: CellReferenceIndex,
        options: &AggregateOptions,
        array_allowed: bool,
        out: &mut Collected,
    ) -> Result<(), CalcResult> {
        if options.skip_nested
            && matches!(
                arg,
                Node::FunctionKind {
                    kind: Function::Subtotal | Function::Aggregate,
                    ..
                }
            )
        {
            return Ok(());
        }
        match self.evaluate_node_with_reference(arg, cell) {
            CalcResult::Range { left, right } => {
                if left.sheet != right.sheet {
                    return Err(CalcResult::new_error(
                        Error::VALUE,
                        cell,
                        "Ranges are in different sheets".to_string(),
                    ));
                }
                let sheet = left.sheet;
                let mut row2 = right.row;
                let mut column2 = right.column;
                // Whole columns / rows: only the used part of the sheet.
                if left.row == 1 && right.row == LAST_ROW {
                    row2 = match self.cached_dimension(sheet) {
                        Ok(d) => d.max_row,
                        Err(m) => return Err(CalcResult::new_error(Error::ERROR, cell, m)),
                    };
                }
                if left.column == 1 && right.column == LAST_COLUMN {
                    column2 = match self.cached_dimension(sheet) {
                        Ok(d) => d.max_column,
                        Err(m) => return Err(CalcResult::new_error(Error::ERROR, cell, m)),
                    };
                }
                let hidden_rows: HashSet<i32> = if options.skip_hidden {
                    match self.workbook.worksheet(sheet) {
                        Ok(ws) => ws.rows.iter().filter(|r| r.hidden).map(|r| r.r).collect(),
                        Err(m) => return Err(CalcResult::new_error(Error::ERROR, cell, m)),
                    }
                } else {
                    HashSet::new()
                };
                for row in left.row..=row2 {
                    if hidden_rows.contains(&row) {
                        continue;
                    }
                    for column in left.column..=column2 {
                        if options.skip_nested && self.cell_is_subtotal_or_aggregate(sheet, row, column) {
                            continue;
                        }
                        let value = self.evaluate_cell(CellReferenceIndex { sheet, row, column });
                        self.aggregate_push(value, options, out)?;
                    }
                }
                Ok(())
            }
            CalcResult::Array(array) => {
                if !array_allowed {
                    return Err(CalcResult::new_error(
                        Error::VALUE,
                        cell,
                        "AGGREGATE functions 1-13 take references, not arrays".to_string(),
                    ));
                }
                for node in array.iter().flatten() {
                    match node {
                        ArrayNode::Number(n) => out.numbers.push(*n),
                        ArrayNode::Error(error) => {
                            if !options.skip_errors {
                                let message = error.to_localized_error_string(self.language);
                                return Err(CalcResult::new_error(error.clone(), cell, message));
                            }
                        }
                        ArrayNode::Boolean(_) | ArrayNode::String(_) => out.others += 1,
                        ArrayNode::Empty => {}
                    }
                }
                Ok(())
            }
            CalcResult::Lambda(_) => Err(CalcResult::new_error(
                Error::VALUE,
                cell,
                "AGGREGATE does not take a LAMBDA".to_string(),
            )),
            value => self.aggregate_push(value, options, out),
        }
    }

    /// One cell or scalar value. An error is skipped (options 2, 3, 6, 7) or
    /// kept: COUNTA counts it, COUNT ignores it, the other functions return
    /// it (`aggregate_reduce`).
    fn aggregate_push(
        &self,
        value: CalcResult,
        options: &AggregateOptions,
        out: &mut Collected,
    ) -> Result<(), CalcResult> {
        match value {
            CalcResult::Number(n) => out.numbers.push(n),
            CalcResult::String(_) | CalcResult::Boolean(_) => out.others += 1,
            error @ CalcResult::Error { .. } => {
                if !options.skip_errors {
                    out.others += 1;
                    if out.first_error.is_none() {
                        out.first_error = Some(error);
                    }
                }
            }
            CalcResult::EmptyCell | CalcResult::EmptyArg => {}
            CalcResult::Range { .. } | CalcResult::Array(_) | CalcResult::Lambda(_) => {}
        }
        Ok(())
    }

    fn cell_is_subtotal_or_aggregate(&self, sheet: u32, row: i32, column: i32) -> bool {
        let formula = self
            .workbook
            .worksheets
            .get(sheet as usize)
            .and_then(|ws| ws.sheet_data.get(&row))
            .and_then(|r| r.get(&column))
            .and_then(|c| c.get_formula());
        match formula {
            Some(f) => matches!(
                self.parsed_formulas
                    .get(sheet as usize)
                    .and_then(|p| p.get(f as usize)),
                Some((
                    Node::FunctionKind {
                        kind: Function::Subtotal | Function::Aggregate,
                        ..
                    },
                    _
                ))
            ),
            None => false,
        }
    }
}

fn div0(cell: CellReferenceIndex) -> CalcResult {
    CalcResult::new_error(Error::DIV, cell, "Divide by 0".to_string())
}

fn mean(values: &[f64]) -> f64 {
    values.iter().sum::<f64>() / values.len() as f64
}

fn sum_squared_deviations(values: &[f64]) -> f64 {
    let m = mean(values);
    values.iter().map(|v| (v - m) * (v - m)).sum()
}

/// Functions 1-13 over the collected values.
fn aggregate_reduce(function_num: i32, collected: Collected, cell: CellReferenceIndex) -> CalcResult {
    if function_num != 2 && function_num != 3 {
        if let Some(error) = collected.first_error {
            return error;
        }
    }
    let values = collected.numbers;
    let n = values.len();
    match function_num {
        // AVERAGE
        1 => {
            if n == 0 {
                return div0(cell);
            }
            CalcResult::Number(mean(&values))
        }
        // COUNT
        2 => CalcResult::Number(n as f64),
        // COUNTA
        3 => CalcResult::Number((n + collected.others) as f64),
        // MAX
        4 => CalcResult::Number(values.iter().copied().fold(None, |m: Option<f64>, v| Some(m.map_or(v, |m| m.max(v)))).unwrap_or(0.0)),
        // MIN
        5 => CalcResult::Number(values.iter().copied().fold(None, |m: Option<f64>, v| Some(m.map_or(v, |m| m.min(v)))).unwrap_or(0.0)),
        // PRODUCT
        6 => CalcResult::Number(if n == 0 { 0.0 } else { values.iter().product() }),
        // STDEV.S
        7 => {
            if n < 2 {
                return div0(cell);
            }
            CalcResult::Number((sum_squared_deviations(&values) / (n - 1) as f64).sqrt())
        }
        // STDEV.P
        8 => {
            if n == 0 {
                return div0(cell);
            }
            CalcResult::Number((sum_squared_deviations(&values) / n as f64).sqrt())
        }
        // SUM
        9 => CalcResult::Number(values.iter().sum()),
        // VAR.S
        10 => {
            if n < 2 {
                return div0(cell);
            }
            CalcResult::Number(sum_squared_deviations(&values) / (n - 1) as f64)
        }
        // VAR.P
        11 => {
            if n == 0 {
                return div0(cell);
            }
            CalcResult::Number(sum_squared_deviations(&values) / n as f64)
        }
        // MEDIAN
        12 => {
            if n == 0 {
                return CalcResult::new_error(Error::NUM, cell, "No numbers".to_string());
            }
            let mut sorted = values;
            sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            let mid = n / 2;
            CalcResult::Number(if n % 2 == 1 {
                sorted[mid]
            } else {
                (sorted[mid - 1] + sorted[mid]) / 2.0
            })
        }
        // MODE.SNGL: the most frequent value; on a tie, the one seen first.
        13 => {
            // value bits -> (count, first position)
            let mut counts: HashMap<u64, (usize, usize)> = HashMap::new();
            for (i, v) in values.iter().enumerate() {
                counts.entry(v.to_bits()).or_insert((0, i)).0 += 1;
            }
            let best = counts
                .iter()
                .filter(|(_, (count, _))| *count > 1)
                .max_by(|a, b| a.1 .0.cmp(&b.1 .0).then(b.1 .1.cmp(&a.1 .1)));
            match best {
                Some((bits, _)) => CalcResult::Number(f64::from_bits(*bits)),
                None => CalcResult::new_error(Error::NA, cell, "No repeated value".to_string()),
            }
        }
        _ => CalcResult::new_error(Error::VALUE, cell, "Invalid AGGREGATE function".to_string()),
    }
}

/// Interpolates the 0-based `rank` in a sorted, non-empty slice.
fn interpolate(sorted: &[f64], rank: f64) -> f64 {
    let low = rank.floor() as usize;
    let high = (low + 1).min(sorted.len() - 1);
    sorted[low] + (rank - rank.floor()) * (sorted[high] - sorted[low])
}

/// Functions 14-19: LARGE, SMALL, PERCENTILE.INC, QUARTILE.INC,
/// PERCENTILE.EXC, QUARTILE.EXC over the sorted values.
fn aggregate_ranked(function_num: i32, sorted: &[f64], k: f64, cell: CellReferenceIndex) -> CalcResult {
    let num_error = |m: &str| CalcResult::new_error(Error::NUM, cell, m.to_string());
    let n = sorted.len();
    match function_num {
        // LARGE / SMALL
        14 | 15 => {
            let k = k.trunc();
            if k < 1.0 || k as usize > n {
                return num_error("k is out of range");
            }
            let k = k as usize;
            CalcResult::Number(if function_num == 14 { sorted[n - k] } else { sorted[k - 1] })
        }
        // PERCENTILE.INC / QUARTILE.INC
        16 | 17 => {
            let p = if function_num == 16 {
                k
            } else {
                let q = k.floor();
                if !(0.0..=4.0).contains(&q) {
                    return num_error("quart must be 0 to 4");
                }
                q / 4.0
            };
            if n == 0 || !(0.0..=1.0).contains(&p) {
                return num_error("k is out of range");
            }
            CalcResult::Number(interpolate(sorted, p * (n - 1) as f64))
        }
        // PERCENTILE.EXC / QUARTILE.EXC
        18 | 19 => {
            let p = if function_num == 18 {
                k
            } else {
                let q = k.floor();
                if !(1.0..=3.0).contains(&q) {
                    return num_error("quart must be 1 to 3");
                }
                q / 4.0
            };
            let nf = n as f64;
            if n == 0 || p <= 0.0 || p >= 1.0 || p < 1.0 / (nf + 1.0) || p > nf / (nf + 1.0) {
                return num_error("k is out of range");
            }
            CalcResult::Number(interpolate(sorted, p * (nf + 1.0) - 1.0))
        }
        _ => CalcResult::new_error(Error::VALUE, cell, "Invalid AGGREGATE function".to_string()),
    }
}
