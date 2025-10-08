use chrono::{NaiveDate, NaiveDateTime};
use pyo3::prelude::*;
use pyo3::types::PyDateTime;

#[derive(Debug, Clone)]
pub enum CellValue {
    Empty,
    String(String),
    Number(f64),
    Bool(bool),
    Date(NaiveDateTime),
}

impl CellValue {
    /// Convert from Python object (used by Dict API)
    pub fn from_py(_py: Python, value: &Bound<PyAny>) -> PyResult<Self> {
        if value.is_none() {
            return Ok(CellValue::Empty);
        }

        if let Ok(s) = value.extract::<&str>() {
            return Ok(CellValue::String(s.to_string()));
        }

        if let Ok(i) = value.extract::<i64>() {
            return Ok(CellValue::Number(i as f64));
        }

        if let Ok(f) = value.extract::<f64>() {
            return Ok(CellValue::Number(f));
        }

        if let Ok(b) = value.extract::<bool>() {
            return Ok(CellValue::Bool(b));
        }

        if let Ok(dt) = value.downcast::<PyDateTime>() {
            use pyo3::types::{PyDateAccess, PyTimeAccess};
            let datetime = NaiveDate::from_ymd_opt(
                dt.get_year(),
                dt.get_month() as u32,
                dt.get_day() as u32,
            )
            .and_then(|date| {
                date.and_hms_opt(
                    dt.get_hour() as u32,
                    dt.get_minute() as u32,
                    dt.get_second() as u32,
                )
            })
            .ok_or_else(|| {
                PyErr::new::<pyo3::exceptions::PyValueError, _>("Invalid datetime")
            })?;

            return Ok(CellValue::Date(datetime));
        }

        Ok(CellValue::String(value.str()?.to_str()?.to_string()))
    }
}

#[derive(Debug, Clone)]
pub struct SheetData {
    pub name: String,
    pub columns: Vec<(String, Vec<CellValue>)>,
}

impl SheetData {
    pub fn validate(&self) -> Result<(), String> {
        if self.name.len() > 31 {
            return Err(format!("Sheet name '{}' exceeds 31 chars", self.name));
        }

        if self.name.chars().any(|c| "[]':*?/\\".contains(c)) {
            return Err(format!("Sheet name '{}' contains invalid chars", self.name));
        }

        if self.columns.is_empty() {
            return Ok(());
        }

        let expected_len = self.columns[0].1.len();
        for (name, col) in &self.columns {
            if col.len() != expected_len {
                return Err(format!(
                    "Column '{}' has {} rows, expected {}",
                    name,
                    col.len(),
                    expected_len
                ));
            }
        }

        Ok(())
    }

    pub fn num_rows(&self) -> usize {
        self.columns.first().map(|(_, col)| col.len()).unwrap_or(0)
    }

    pub fn num_cols(&self) -> usize {
        self.columns.len()
    }
}

#[derive(Debug)]
pub enum WriteError {
    Io(std::io::Error),
    Validation(String),
}

impl std::fmt::Display for WriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match self {
            WriteError::Io(e) => write!(f, "IO error: {}", e),
            WriteError::Validation(e) => write!(f, "Validation error: {}", e),
        }
    }
}

impl From<std::io::Error> for WriteError {
    fn from(e: std::io::Error) -> Self {
        WriteError::Io(e)
    }
}