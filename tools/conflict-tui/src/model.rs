//! 与 C# FileDiffDto 对齐的 serde 数据模型。
//! schema 见 NumDesTools.Core/ConflictResolver/FileDiffDto.cs（camelCase + JsonStringEnumConverter）。

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FileDiffDto {
    pub ours_path: String,
    pub theirs_path: String,
    #[serde(default)]
    pub ours_label: Option<String>,
    #[serde(default)]
    pub theirs_label: Option<String>,
    pub sheets: Vec<SheetDiffDto>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetDiffDto {
    pub sheet_name: String,
    pub all_columns: Vec<String>,
    pub type_row: std::collections::HashMap<String, String>,
    pub label_row: std::collections::HashMap<String, String>,
    pub rows: Vec<RowConflictDto>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RowConflictDto {
    pub sheet_name: String,
    pub row_key: String,
    pub diff_type: RowDiffType,
    pub origin: RowOrigin,
    pub ours_row_index: i32,
    pub theirs_row_index: i32,
    pub all_columns: Vec<String>,
    pub ours_full_row: Option<std::collections::HashMap<String, Option<String>>>,
    pub theirs_full_row: Option<std::collections::HashMap<String, Option<String>>>,
    pub cells: Vec<CellConflictDto>,
    pub row_choice: ConflictChoice,
    pub row_choice_explicit: bool,
    pub ai_suggestion: String,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellConflictDto {
    pub col_name: String,
    pub ours_value: Option<String>,
    pub theirs_value: Option<String>,
    pub choice: ConflictChoice,
    pub is_explicit: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Serialize)]
pub enum ConflictChoice {
    Ours,
    Theirs,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Serialize)]
pub enum RowDiffType {
    Modified,
    OnlyOurs,
    OnlyTheirs,
    Same,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Serialize)]
pub enum RowOrigin {
    Unknown,
    AddedByOurs,
    DeletedByTheirs,
    AddedByTheirs,
    DeletedByOurs,
}

/// Rust TUI 回传的用户选择（只传变化量，不回传完整 FileDiff）。
#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SelectionResultDto {
    pub confirmed: bool,
    pub selections: Vec<SelectionDto>,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SelectionDto {
    pub sheet_name: String,
    pub row_key: String,
    pub col_name: Option<String>, // null = 整行（OnlyOurs/OnlyTheirs）
    pub choice: ConflictChoice,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn roundtrip_file_diff_dto() {
        let json = r#"{
            "oursPath": "C:\\a.xlsx",
            "theirsPath": "C:\\b.xlsx",
            "sheets": [{
                "sheetName": "Sheet1",
                "allColumns": ["id", "hp"],
                "typeRow": {"id": "int"},
                "labelRow": {"id": "编号"},
                "rows": [{
                    "sheetName": "Sheet1",
                    "rowKey": "1001",
                    "diffType": "Modified",
                    "origin": "Unknown",
                    "oursRowIndex": 5,
                    "theirsRowIndex": 5,
                    "allColumns": ["id", "hp"],
                    "oursFullRow": {"id": "1001", "hp": "100"},
                    "theirsFullRow": {"id": "1001", "hp": "120"},
                    "cells": [{
                        "colName": "hp",
                        "oursValue": "100",
                        "theirsValue": "120",
                        "choice": "Ours",
                        "isExplicit": false
                    }],
                    "rowChoice": "Ours",
                    "rowChoiceExplicit": false,
                    "aiSuggestion": ""
                }]
            }]
        }"#;
        let dto: FileDiffDto = serde_json::from_str(json).unwrap();
        assert_eq!(dto.sheets.len(), 1);
        assert_eq!(dto.sheets[0].rows[0].cells[0].col_name, "hp");
        assert_eq!(
            dto.sheets[0].rows[0].cells[0].choice,
            ConflictChoice::Ours
        );
        let serialized = serde_json::to_string(&dto).unwrap();
        let deserialized: FileDiffDto = serde_json::from_str(&serialized).unwrap();
        assert_eq!(deserialized.sheets[0].rows[0].row_key, "1001");
    }

    #[test]
    fn selection_result_roundtrip() {
        let json = r#"{
            "confirmed": true,
            "selections": [
                { "sheetName": "Sheet1", "rowKey": "1001", "colName": "hp", "choice": "Theirs" },
                { "sheetName": "Sheet1", "rowKey": "1002", "colName": null, "choice": "Ours" }
            ]
        }"#;
        let dto: SelectionResultDto = serde_json::from_str(json).unwrap();
        assert!(dto.confirmed);
        assert_eq!(dto.selections.len(), 2);
        assert_eq!(dto.selections[1].col_name, None);
    }
}
