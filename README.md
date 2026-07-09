# table-transform

Convert tabular data between markdown, CSV, TSV, Excel (xlsx), JSON, and plain text. Includes a CLI and a TypeScript library for reading, writing, and transforming tables.

[![npm Package Version](https://img.shields.io/npm/v/table-transform)](https://www.npmjs.com/package/table-transform)

## Features

- Convert between **markdown**, **csv**, **tsv**, **xlsx**, **json**, and **txt**
- Read multiple markdown tables from one file
- Read/write **txt** as SQLite `.mode table` ASCII box tables (auto-padded columns)
- Read multiple Excel sheets, with optional cell range
- Auto-detect CSV/TSV separator (`,` or `\t`)
- Trim whitespace, empty rows, and empty columns by default (can be disabled)
- Export JSON as array of objects or 2D array of values
- TypeScript support with typed return values per format

## Installation

Install as a library in your project:

```bash
npm install table-transform
```

Install the CLI globally:

```bash
npm install -g table-transform
```

Or run without installing:

```bash
npx -y table-transform --help
npx -y table-transform source.xlsx export.csv
```

Use `-y` to skip the install prompt in non-interactive environments.

You can also install with [pnpm](https://pnpm.io/), [yarn](https://yarnpkg.com/), or [slnpm](https://github.com/beenotung/slnpm).

## CLI Usage

```
Usage:
  table-transform [options]

Alias: convert-table, extract-table, export-table

Options:
  -i, --input <file>     Input file (path)
  -o, --output <file>    Output file (path or /dev/stdout)
  -h, --help             Show help message
  -v, --version          Show version

Options to disable trimming (default is enabled):
  --no-trim-string       Preserve leading/trailing whitespace characters in string values
  --no-trim-rows         Preserve leading/trailing empty rows
  --no-trim-cols         Preserve leading/trailing empty columns

Options for csv/tsv files:
  -s, --separator <char>     Example: '|' (default auto detect ',' or '\t')
  --input-separator <char>   Default same as --separator
  --output-separator <char>  Default same as --separator

Options for json files:
  --array       As 2D array of values
  --object      As array of key-value objects (default)

Options for console output:
  -f, --format <format>  Output format (default: markdown)
  -n, --name <mode>      Display table name or not (default: auto)
  --show-name                     Alias for "--name always"
  --hide-name                     Alias for "--name never"

Supported formats:
  - md, markdown
  - csv, tsv
  - txt (ASCII box format from SQLite `.mode table`)
  - xlsx
  - json

Display table name modes:
  auto   : display table name only if there are multiple tables
  always : always display table name
  never  : never display table name

When a file contains multiple tables or sheets, output files are split as {basename}-{sheet-name}.{ext}.
```

Run `table-transform --help` for example commands.

## Library Usage Example

For more examples, see: [src/test.ts](./src/test.ts)

### Convert between file formats

```typescript
import { convert_file } from 'table-transform'

convert_file({
  source_file: 'res/roster.md',
  dest_file: 'res/roster.csv',
})

convert_file({
  source_file: 'res/roster.csv',
  dest_file: 'res/roster.json',
  json_format: 'object', // default
})
```

### Read and write

```typescript
import { read_file, write_file } from 'table-transform'

let sheets = read_file({ file: 'res/roster.md' })
// sheets[0].name, sheets[0].rows

write_file({
  file: 'res/roster.xlsx',
  sheets,
})
```

`read_file` returns `SheetData<string>[]` for markdown/csv/tsv/txt, and `SheetData<CellValue>[]` for xlsx/json.

### Read markdown with multiple tables

```typescript
import { read_md_file } from 'table-transform'

let sheets = read_md_file({ file: 'res/multi-table.md' })
// one SheetData per markdown table in the file
```

### Read SQLite-style txt tables

`.txt` uses the ASCII box format from SQLite `.mode table`: auto-padded columns, centered headers, right-aligned numbers, left-aligned text. Newlines in cells are escaped as `\n`. Multiple tables in one file are supported on read.

```typescript
import { read_txt_file } from 'table-transform'

let sheets = read_txt_file({ file: 'res/roster.txt' })
// one SheetData per table in the file
```

### Read Excel with sheet range

```typescript
import { read_xlsx_file } from 'table-transform'

let sheets = read_xlsx_file({
  file: 'res/Attendance.xlsx',
  sheets: [{ name: 'Sheet1', range: 'A14:E26' }],
})
```

### Convert rows to objects

```typescript
import { read_csv_file, rows_to_objects } from 'table-transform'

let rows = read_csv_file({ file: 'res/roster.csv' }).rows
let objects = rows_to_objects(rows)
// [{ No: '1', 'Chi Name': '陳小明', ... }, ...]
```

### JSON formats

```typescript
import { write_json_file } from 'table-transform'

// array of objects (default)
write_json_file({ file: 'out.json', rows, format: 'object' })

// 2D array
write_json_file({ file: 'out.json', rows, format: 'array' })
```

## Typescript Signature

Shared read options are grouped as `ExtraReadFileOptions` below for readability. It is not exported from the package.

```typescript
export type CellValue = string | number | boolean | Date | null

export type SheetData<T extends CellValue> = {
  name: string
  rows: T[][]
}

export type ShowNameMode = 'auto' | 'always' | 'never'

export type SheetInfo = {
  name: string
  range?: string
}

type ExtraReadFileOptions = {
  /** for csv/tsv, default auto detect ',' or '\t' */
  separator?: string
  /** trim string value, default true */
  trim_string?: boolean
  /** trim empty leading/trailing rows, default true */
  trim_rows?: boolean
  /** trim empty leading/trailing cols, default true */
  trim_cols?: boolean
}

export function read_file(
  args: {
    file:
      | `${string}.md`
      | `${string}.markdown`
      | `${string}.csv`
      | `${string}.tsv`
      | `${string}.txt`
  } & ExtraReadFileOptions,
): SheetData<string>[]
export function read_file(
  args: {
    file: `${string}.xlsx` | `${string}.json` | string
  } & ExtraReadFileOptions,
): SheetData<CellValue>[]
export function read_file(
  args: { file: string } & ExtraReadFileOptions,
): SheetData<CellValue>[]

export function write_file(args: {
  file: string
  sheets: SheetData<CellValue>[]
  /** only used when file is /dev/stdout, default 'auto' */
  show_name?: ShowNameMode
  /** for csv/tsv files */
  separator?: string
  /** for json files, default 'object' */
  json_format?: 'object' | 'array'
}): void

export function convert_file(
  args: {
    source_file: string
    dest_file: string
    /** for output json files, default 'object' */
    json_format?: 'object' | 'array'
    /** default same as separator */
    input_separator?: string
    /** default same as separator */
    output_separator?: string
  } & ExtraReadFileOptions,
): void

export function read_xlsx_file(
  args: {
    file: string
    sheets?: string[] | SheetInfo[]
  } & ExtraReadFileOptions,
): SheetData<CellValue>[]

export function read_csv_file(
  args: {
    file: string
  } & ExtraReadFileOptions,
): SheetData<string> & { separator: string }

export function read_md_file(
  args: { file: string } & ExtraReadFileOptions,
): SheetData<string>[]

export function read_txt_file(args: { file: string }): SheetData<string>[]

export function read_json_file(
  args: {
    file: string
    format?: 'array' | 'object'
  } & ExtraReadFileOptions,
): SheetData<CellValue>

export function write_csv_file(args: {
  file: string
  rows: CellValue[][]
  separator?: string
}): void

export function write_xlsx_file(args: {
  file: string
  sheets: SheetData<CellValue>[]
}): void

export function write_md_file(args: { file: string; rows: CellValue[][] }): void

export function write_txt_file(args: {
  file: string
  rows: CellValue[][]
}): void

export function write_json_file(args: {
  file: string
  rows: CellValue[][]
  /** @default 'object' */
  format?: 'array' | 'object'
}): void

export function infer_csv_separator(text: string): {
  separator: ',' | '\t'
  rows: string[][]
}

export function rows_to_objects<T>(rows: T[][]): Record<string, T>[]

export function objects_to_rows<T>(
  objects: Record<string, T>[],
): (T | string | null)[][]
```

## License

This project is licensed with [BSD-2-Clause](./LICENSE)

This is free, libre, and open-source software. It comes down to four essential freedoms [[ref]](https://seirdy.one/2021/01/27/whatsapp-and-the-domestication-of-users.html#fnref:2):

- The freedom to run the program as you wish, for any purpose
- The freedom to study how the program works, and change it so it does your computing as you wish
- The freedom to redistribute copies so you can help others
- The freedom to distribute copies of your modified versions to others
