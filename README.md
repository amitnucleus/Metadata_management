# Metadata Management

A standalone Next.js app for uploading, editing, and downloading data mapping specification files.

## Features

- Upload **XLSX / XLS / CSV / JSON** files
- Editable spreadsheet table with inline cell editing
- Column validation rules (enum, integer, regex) with auto-detection by column name
- Manually pin columns to specific rules
- Export back as **XLSX** or **JSON**

## Getting Started

```bash
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

## Supported File Types

| Format | Upload | Download |
|--------|--------|----------|
| XLSX / XLS / CSV | Yes | Yes |
| JSON | Yes | Yes |

## Column Validation Rules

Rules auto-match columns by name pattern. Supported rule types:

| Rule | Type | Example columns |
|------|------|-----------------|
| File Format | enum | `file_format`, `fileformat` |
| File Encoding | enum | `encoding`, `file_encoding` |
| Column Data Type | enum | `data_type`, `column_data_type` |
| Action Type | enum | `action_type` |
| Feed / File Type | enum | `feed_type`, `feed_file_type` |
| Mandatory | enum | `mandatory`, `required` |
| Derived Logic | enum | `derived_logic`, `transformation_logic` |
| Integer / Position | integer | `position`, `index`, `sequence` |
| Column Format | regex | `column_format`, `format_string` |
