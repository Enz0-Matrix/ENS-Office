//! ENS-Office: incremental recalculation.
//!
//! A full evaluation (`Model::evaluate`) records, for every formula, what it
//! read: single positions (every read goes through `evaluate_cell`) and
//! rectangles (ranges read straight from the sheet, the empty remainder of
//! whole-column reads, lookup vectors answered from the per-pass index). It
//! also takes a fingerprint of every cell: the value of a constant, the text
//! of a formula.
//!
//! `Model::evaluate_incremental` then finds what changed by comparing the
//! fingerprints, however the change was made, marks the formulas that depend
//! on it, transitively, and runs only those. Every other formula keeps its
//! stored value. Whenever the shortcut could be wrong it falls back to the
//! full evaluation:
//!
//! * no graph yet (nothing was evaluated in full since load),
//! * arrays that spill anywhere, or CSE arrays (they write cells other than
//!   their own); a dynamic formula whose result is a single value (INDEX,
//!   XLOOKUP...) is an ordinary formula here, and if it starts to spill the
//!   evaluation is redone in full,
//! * anything but cell contents changed: sheets, defined names, tables,
//!   locale, hidden rows or columns (SUBTOTAL),
//! * a large share of the formulas would have to run anyway,
//! * a circular reference among the formulas that ran.
//!
//! Volatile functions (NOW, TODAY, RAND..., CELL, INFO) run on every
//! evaluation, as in Excel.
//!
//! Cost. While formulas run, a read is one push to a flat list (`edges`); the
//! reverse index, position → readers, is built from it only when an
//! incremental evaluation first needs it. Old entries of a formula that ran
//! again are left in the reverse index: they can only make a formula run
//! that did not need to, never keep one from running.

use std::collections::{HashMap, HashSet};
use std::hash::{Hash, Hasher};

use crate::evaluation::{CellKey, FormulaCell, KeyBuild};
use crate::expressions::token::Error;
use crate::expressions::types::CellReferenceIndex;
use crate::model::Model;
use crate::types::{ArrayKind, Cell, FormulaValue};

/// `(sheet, row1, column1, row2, column2)`, inclusive.
pub(crate) type Rect = (u32, i32, i32, i32, i32);

fn rect_contains(rect: &Rect, (sheet, row, column): CellKey) -> bool {
    rect.0 == sheet && (rect.1..=rect.3).contains(&row) && (rect.2..=rect.4).contains(&column)
}

#[derive(Clone, Copy)]
enum Read {
    Cell(CellKey),
    Rect(Rect),
}

/// The dependency graph and the fingerprints taken after the last evaluation.
#[derive(Default)]
pub(crate) struct DepGraph {
    /// Set by a full evaluation, cleared when the graph cannot be trusted.
    pub(crate) valid: bool,
    /// Reads are recorded while this is set (during every evaluation).
    pub(crate) recording: bool,
    /// `(reader, run, read)`, in the order the reads were made. A formula
    /// that runs again gets a new run number; only its last run counts.
    edges: Vec<(CellKey, u32, Read)>,
    /// Formula → number of its last run.
    last_run: HashMap<CellKey, u32, KeyBuild>,
    run_counter: u32,
    /// The formula the last read was made for, and its run: most reads come
    /// in a row from the same formula, so most need no lookup.
    current: Option<(CellKey, u32)>,
    /// Position → formulas that read it, built from `edges` on demand.
    rev_cells: HashMap<CellKey, Vec<CellKey>, KeyBuild>,
    /// Rectangle → formulas that read it.
    rev_rects: HashMap<Rect, Vec<CellKey>, KeyBuild>,
    /// Formulas that called a volatile function in their last run.
    pub(crate) volatile: HashSet<CellKey, KeyBuild>,
    /// Formulas found on a cycle in their last run. A cell marked circular is
    /// not run again within the pass, so what it reads is not known: it runs
    /// on every incremental evaluation, like a volatile one.
    pub(crate) circular: HashSet<CellKey, KeyBuild>,
    /// Fingerprint of every cell, in the order the sheets give their cells
    /// (sheet, row, column): compared by merging, without hashing.
    fingerprints: Vec<(CellKey, u64)>,
    /// Fingerprint of everything else results depend on.
    structure: u64,
    /// When set, only these formulas run; every other one is taken as
    /// evaluated (its stored value is read).
    pub(crate) only: Option<HashSet<CellKey, KeyBuild>>,
    /// For tests and the benchmark: how the last evaluation was done.
    last: LastEvaluation,
}

/// How the last evaluation was done.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub enum LastEvaluation {
    #[default]
    Full,
    /// Number of formulas run.
    Incremental(usize),
}

impl DepGraph {
    /// Forgets everything recorded: a full evaluation starts.
    pub(crate) fn clear_recording(&mut self) {
        self.edges.clear();
        self.last_run.clear();
        self.current = None;
        self.rev_cells.clear();
        self.rev_rects.clear();
        self.volatile.clear();
        self.circular.clear();
    }

    /// Moves the recorded reads of the last run of each formula into the
    /// reverse index.
    fn index_edges(&mut self) {
        let edges = std::mem::take(&mut self.edges);
        for (reader, run, read) in edges {
            if self.last_run.get(&reader) != Some(&run) {
                continue;
            }
            let readers = match read {
                Read::Cell(cell) => self.rev_cells.entry(cell).or_default(),
                Read::Rect(rect) => self.rev_rects.entry(rect).or_default(),
            };
            // A formula that runs again reads mostly the same positions: do
            // not let the lists grow with every edit.
            let known = if readers.len() <= 32 {
                readers.contains(&reader)
            } else {
                readers.last() == Some(&reader)
            };
            if !known {
                readers.push(reader);
            }
        }
        self.current = None;
    }

    fn push(&mut self, top: CellKey, read: Read) {
        let run = match self.current {
            Some((key, run)) if key == top => run,
            _ => {
                let run = match self.last_run.get(&top) {
                    Some(run) => *run,
                    None => {
                        self.run_counter += 1;
                        self.last_run.insert(top, self.run_counter);
                        self.run_counter
                    }
                };
                self.current = Some((top, run));
                run
            }
        };
        self.edges.push((top, run, read));
    }
}

#[inline]
fn mix(a: u64, b: u64) -> u64 {
    (a.rotate_left(5) ^ b).wrapping_mul(0x517c_c1b7_2722_0a95)
}

fn hash_of<T: Hash + ?Sized>(value: &T) -> u64 {
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    value.hash(&mut hasher);
    hasher.finish()
}

/// A dynamic formula blocked from spilling (`#SPILL!`) depends on the cells
/// in its way, which it does not read through `evaluate_cell`.
fn is_spill_error(value: &FormulaValue) -> bool {
    matches!(value, FormulaValue::Error { ei: Error::SPILL, .. })
}

/// What a scan of the cells found.
struct Scan {
    changed: Vec<CellKey>,
    formula_count: usize,
    /// Formulas still `Unevaluated` (entered again with the same text).
    unevaluated: Vec<CellKey>,
    has_arrays: bool,
}

impl Model<'_> {
    #[inline]
    fn top_formula(&self) -> Option<CellKey> {
        self.evaluation
            .stack
            .last()
            .map(|top| (top.sheet, top.row, top.column))
    }

    /// Called by the evaluation for every read made by the formula on top of
    /// the stack.
    #[inline]
    pub(crate) fn record_read(&mut self, position: CellReferenceIndex) {
        if !self.deps.recording {
            return;
        }
        if let Some(top) = self.top_formula() {
            let read = Read::Cell((position.sheet, position.row, position.column));
            self.deps.push(top, read);
        }
    }

    /// A rectangle read by the formula on top of the stack.
    pub(crate) fn record_read_rect(&mut self, rect: Rect) {
        if !self.deps.recording {
            return;
        }
        if let Some(top) = self.top_formula() {
            self.deps.push(top, Read::Rect(rect));
        }
    }

    /// The formula on top of the stack called a volatile function.
    pub(crate) fn record_volatile(&mut self) {
        if let Some(top) = self.top_formula() {
            self.deps.volatile.insert(top);
        }
    }

    /// A formula is about to run (again): what it read before no longer counts.
    pub(crate) fn reset_reads(&mut self, formula: CellKey) {
        if self.deps.recording {
            self.deps.run_counter += 1;
            let run = self.deps.run_counter;
            self.deps.last_run.insert(formula, run);
            self.deps.current = Some((formula, run));
            self.deps.volatile.remove(&formula);
            self.deps.circular.remove(&formula);
        }
    }

    /// Everything but cell contents that results depend on.
    fn structure_fingerprint(&self) -> u64 {
        let mut h = std::collections::hash_map::DefaultHasher::new();
        let wb = &self.workbook;
        wb.settings.tz.hash(&mut h);
        wb.settings.locale.hash(&mut h);
        self.get_language().hash(&mut h);
        for name in &wb.defined_names {
            name.name.hash(&mut h);
            name.formula.hash(&mut h);
            name.sheet_id.hash(&mut h);
        }
        let mut tables: Vec<String> = wb.tables.iter().map(|t| format!("{t:?}")).collect();
        tables.sort();
        tables.hash(&mut h);
        wb.worksheets.len().hash(&mut h);
        for ws in &wb.worksheets {
            ws.name.hash(&mut h);
            ws.sheet_id.hash(&mut h);
            format!("{:?}", ws.state).hash(&mut h);
            for row in ws.rows.iter().filter(|r| r.hidden) {
                row.r.hash(&mut h);
            }
            for col in ws.cols.iter().filter(|c| c.hidden) {
                (col.min, col.max).hash(&mut h);
            }
        }
        h.finish()
    }

    /// Compares every cell with its fingerprint and updates the fingerprints.
    /// With `reset`, the fingerprints are taken afresh and nothing is
    /// reported as changed.
    fn scan_cells(&mut self, reset: bool) -> Scan {
        let mut scan = Scan {
            changed: Vec::new(),
            formula_count: 0,
            unevaluated: Vec::new(),
            has_arrays: false,
        };
        let old = std::mem::take(&mut self.deps.fingerprints);
        let mut fingerprints: Vec<(CellKey, u64)> = Vec::with_capacity(old.len() + 16);
        let mut previous = old.iter().peekable();
        for (sheet, ws) in self.workbook.worksheets.iter().enumerate() {
            let sheet = sheet as u32;
            let formula_hashes: Vec<u64> = ws.shared_formulas.iter().map(|f| hash_of(f)).collect();
            let shared_strings = &self.workbook.shared_strings;
            for (row, column, cell) in ws.sheet_data.cells() {
                let position = (sheet, row, column);
                let fingerprint = match &*cell {
                    Cell::EmptyCell { .. } => 1,
                    Cell::BooleanCell { v, .. } => mix(2, *v as u64),
                    Cell::NumberCell { v, .. } => mix(3, v.to_bits()),
                    Cell::ErrorCell { ei, .. } => mix(4, hash_of(&format!("{ei:?}"))),
                    Cell::SharedString { si, .. } => {
                        let length = shared_strings.get(*si as usize).map_or(0, |s| s.len());
                        mix(mix(5, *si as u64), length as u64)
                    }
                    Cell::CellFormula { f, v, .. } => {
                        scan.formula_count += 1;
                        if matches!(v, FormulaValue::Unevaluated) {
                            scan.unevaluated.push(position);
                        }
                        mix(6, formula_hashes.get(*f as usize).copied().unwrap_or(0))
                    }
                    Cell::ArrayFormula { f, r, kind, v, .. }
                        if *r == (1, 1)
                            && matches!(kind, ArrayKind::Dynamic)
                            && !is_spill_error(v) =>
                    {
                        // A dynamic formula that does not spill.
                        scan.formula_count += 1;
                        if matches!(v, FormulaValue::Unevaluated) {
                            scan.unevaluated.push(position);
                        }
                        mix(7, formula_hashes.get(*f as usize).copied().unwrap_or(0))
                    }
                    Cell::ArrayFormula { .. } | Cell::SpillCell { .. } => {
                        scan.has_arrays = true;
                        8
                    }
                };
                if !reset {
                    // Old cells before this position are gone.
                    while let Some((key, _)) = previous.next_if(|(key, _)| *key < position) {
                        scan.changed.push(*key);
                    }
                    match previous.next_if(|(key, _)| *key == position) {
                        Some((_, old_fingerprint)) if *old_fingerprint == fingerprint => {}
                        _ => scan.changed.push(position),
                    }
                }
                fingerprints.push((position, fingerprint));
            }
        }
        if !reset {
            scan.changed.extend(previous.map(|(key, _)| *key));
        }
        self.deps.fingerprints = fingerprints;
        scan
    }

    fn is_formula(&self, (sheet, row, column): CellKey) -> bool {
        self.fetch_cell(CellReferenceIndex { sheet, row, column })
            .as_deref()
            .and_then(FormulaCell::of)
            .is_some()
    }

    /// After a full evaluation: the reads recorded during it become the graph
    /// incremental evaluations work from; the reverse index is built when
    /// first needed.
    pub(crate) fn finish_full_evaluation(&mut self) {
        let scan = self.scan_cells(true);
        self.deps.valid = !scan.has_arrays;
        self.deps.structure = self.structure_fingerprint();
        self.deps.rev_cells.clear();
        self.deps.rev_rects.clear();
        self.deps.last = LastEvaluation::Full;
        if !self.deps.valid {
            self.deps.clear_recording();
        }
    }

    /// Evaluates only the formulas affected by what changed since the last
    /// evaluation, falling back to [`Model::evaluate`] whenever that could
    /// give a different result. See the module documentation.
    pub fn evaluate_incremental(&mut self) {
        if !self.deps.valid || self.structure_fingerprint() != self.deps.structure {
            self.evaluate();
            return;
        }
        let scan = self.scan_cells(false);
        if scan.has_arrays {
            self.evaluate();
            return;
        }
        self.deps.index_edges();

        let mut volatile = std::mem::take(&mut self.deps.volatile);
        let mut circular = std::mem::take(&mut self.deps.circular);
        // Cells that are no longer formulas are neither volatile nor circular.
        volatile.retain(|k| self.is_formula(*k));
        circular.retain(|k| self.is_formula(*k));
        let mut queue = scan.changed;
        queue.extend(scan.unevaluated.iter().copied());
        queue.extend(volatile.iter().copied());
        queue.extend(circular.iter().copied());
        self.deps.volatile = volatile;
        self.deps.circular = circular;

        let limit = (scan.formula_count / 2).max(1000);
        let mut dirty: HashSet<CellKey, KeyBuild> = HashSet::default();
        let mut visited: HashSet<CellKey, KeyBuild> = HashSet::default();
        let rects: Vec<Rect> = self.deps.rev_rects.keys().copied().collect();
        while let Some(position) = queue.pop() {
            if !visited.insert(position) {
                continue;
            }
            if self.is_formula(position) {
                dirty.insert(position);
                if dirty.len() > limit {
                    self.evaluate();
                    return;
                }
            }
            if let Some(readers) = self.deps.rev_cells.get(&position) {
                queue.extend(readers.iter().filter(|r| !visited.contains(*r)));
            }
            for rect in &rects {
                if rect_contains(rect, position) {
                    let readers = &self.deps.rev_rects[rect];
                    queue.extend(readers.iter().filter(|r| !visited.contains(*r)));
                }
            }
        }

        let mut roots: Vec<CellKey> = dirty.iter().copied().collect();
        roots.sort_unstable();
        let run = roots.len();
        // With a cycle the result depends on the order the cells are visited
        // in (cells on the loop either fail or read an error): redo it in full
        // so that it is the same as a full evaluation.
        if self.run_incremental_pass(&roots, dirty).is_err()
            || self.spilled(&roots)
            || roots.iter().any(|f| self.deps.circular.contains(f))
        {
            self.evaluate();
            return;
        }
        self.deps.index_edges();
        self.deps.last = LastEvaluation::Incremental(run);
        self.evaluate_conditional_formatting();
        #[cfg(test)]
        self.verify_incremental();
    }

    /// Whether one of `formulas` now spills or wants to (its result grew
    /// beyond its cell, or is blocked: `#SPILL!`).
    fn spilled(&self, formulas: &[CellKey]) -> bool {
        formulas.iter().any(|&(sheet, row, column)| {
            matches!(
                self.fetch_cell(CellReferenceIndex { sheet, row, column }).as_deref(),
                Some(Cell::ArrayFormula { r, v, .. }) if *r != (1, 1) || is_spill_error(v)
            )
        })
    }

    /// Runs `roots` with every other formula taken as evaluated.
    fn run_incremental_pass(
        &mut self,
        roots: &[CellKey],
        dirty: HashSet<CellKey, KeyBuild>,
    ) -> Result<(), ()> {
        let state = &mut self.evaluation;
        state.cells.clear();
        state.stack.clear();
        state.seen.clear();
        state.seen_ranges.clear();
        state.lookup_cache.clear();
        state.restart = None;
        state.depth = 0;
        state.unwinding = false;
        state.circular.clear();
        state.in_pass = true;
        self.deps.only = Some(dirty);
        self.deps.recording = true;
        self.clear_variable_stack();
        self.clear_lambdas();
        let mut result = Ok(());
        for &(sheet, row, column) in roots {
            let cell = CellReferenceIndex { sheet, row, column };
            self.evaluation.root = Some(cell);
            self.evaluate_cell(cell);
            if self.evaluation.restart.is_some() {
                result = Err(());
                break;
            }
        }
        self.evaluation.restart = None;
        self.evaluation.in_pass = false;
        self.evaluation.root = None;
        self.evaluation.lookup_cache.clear();
        self.deps.only = None;
        self.deps.recording = false;
        if result.is_err() {
            self.deps.valid = false;
        }
        result
    }

    /// How the last evaluation was done (for tests and benchmarks).
    pub fn last_evaluation(&self) -> LastEvaluation {
        self.deps.last
    }

    /// Under test with `ENS_VERIFY_INCREMENTAL=1`: every incremental
    /// evaluation is checked against a full one.
    #[cfg(test)]
    fn verify_incremental(&mut self) {
        if std::env::var("ENS_VERIFY_INCREMENTAL").as_deref() != Ok("1")
            || !self.deps.volatile.is_empty()
        {
            return;
        }
        let snapshot = |model: &Model| -> Vec<(CellKey, String)> {
            let mut values = Vec::new();
            for (sheet, ws) in model.workbook.worksheets.iter().enumerate() {
                for (row, column, cell) in ws.sheet_data.cells() {
                    values.push(((sheet as u32, row, column), format!("{cell:?}")));
                }
            }
            values
        };
        let incremental = snapshot(self);
        let last = self.deps.last;
        self.evaluate();
        let full = snapshot(self);
        if incremental != full {
            let diff: Vec<_> = incremental
                .iter()
                .zip(full.iter())
                .filter(|(a, b)| a != b)
                .take(5)
                .collect();
            panic!("ENS incremental evaluation differs from full: {diff:?}");
        }
        self.deps.last = last;
    }
}
