# table-transform

Convert tabular data between markdown, CSV, TSV, Excel (xlsx), JSON, plain text, and Google Sheets. Includes a CLI and a TypeScript library for reading, writing, and transforming tables.

[![npm Package Version](https://img.shields.io/npm/v/table-transform)](https://www.npmjs.com/package/table-transform)

## Features

- Convert between **markdown**, **csv**, **tsv**, **xlsx**, **json**, and **txt**
- Read and write **Google Sheets** with a service account
- Read multiple markdown tables from one file
- Read/write **txt** as SQLite `.mode table` ASCII box tables (auto-padded columns)
- Read multiple Excel sheets, with optional cell range
- Auto-detect CSV/TSV separator (`,` or `\t`)
- Trim whitespace, empty rows, and empty columns by default (can be disabled)
- Pick or exclude columns by header name
- Export JSON as array of objects or 2D array of values
- TypeScript support with typed return values per format

## Installation

Requires Node.js 18 or newer.

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
  -l, --list             List sheet/table names
  -h, --help             Show help message
  -v, --version          Show version

Options to disable trimming (default is enabled):
  --no-trim-string       Preserve leading/trailing whitespace characters in string values
  --no-trim-rows         Preserve leading/trailing empty rows
  --no-trim-cols         Preserve leading/trailing empty columns

Options for Google Sheets:
  --sa-cred-file <file>      Service account credential json file

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
  --show-name            Alias for "--name always"
  --hide-name            Alias for "--name never"

Options for column selection (default: all columns; comma-separated or repeat flag):
  -c, --column <names>   Include columns by header name
                         Alias: --columns
                                --field, --fields
  -x, --exclude <names>  Exclude columns by header name
                         Alias: --exclude-column, --exclude-columns
                                --exclude-field, --exclude-fields

Supported formats:
  - md, markdown
  - csv, tsv
  - txt (ASCII box format from SQLite `.mode table`)
  - xlsx
  - json

Google Sheets references (use in place of an input/output file):
  gsheets:<spreadsheet-id>                 All tabs
  gsheets:<spreadsheet-id>/<sheet-name>    One tab, by name
  gsheets:<spreadsheet-id>#<sheet-index>   One tab, by zero-based index

Display table name modes:
  auto   : display table name only if there are multiple tables
  always : always display table name
  never  : never display table name

When a file contains multiple tables or sheets, output files are split as {basename}-{sheet-name}.{ext}.
```

List sheet/table names:

```bash
table-transform --list source.xlsx
table-transform multi-table.md -l
```

Pick selected columns by header name (comma-separated or repeat the flag):

```bash
# pick selected columns (instead of all columns)
table-transform --column "Name, Tel" source.csv export.csv
table-transform source.md -c "Name,Tel" export.csv
table-transform source.csv -c Name -c Tel export.csv

# exclude selected columns (include all other columns)
table-transform --exclude "Email,CV" source.csv export.csv
table-transform source.csv -x "Email, CV" export.csv
```

Run `table-transform --help` for more example commands.

### Google Sheets

Use a `gsheets:` reference wherever a file path is accepted. The spreadsheet id is the long token in the sheet URL, between `/d/` and `/edit`:

```
https://docs.google.com/spreadsheets/d/1AbC.....XyZ/edit#gid=0
                                       ^^^^^^^^^^^^^
```

| Reference | Selects |
| --- | --- |
| `gsheets:<spreadsheet-id>` | All tabs, like the sheets of an xlsx workbook |
| `gsheets:<spreadsheet-id>/<sheet-name>` | One tab, by name |
| `gsheets:<spreadsheet-id>#<sheet-index>` | One tab, by zero-based index |

A tab name is taken verbatim, so `gsheets:1AbC/2024` means the tab **named** `2024`, and only `gsheets:1AbC#3` means "the fourth tab". A name may itself contain `#`, since the selector is whichever of `/` or `#` comes first.

**Setup:**

1. In the [Google Cloud console](https://console.cloud.google.com/), create a service account and enable the **Google Sheets API** for its project.
2. Create a JSON key for that service account and download it.
3. **Share the spreadsheet with the service account's email address** (the `client_email` in the JSON key), as Editor for writing or Viewer for reading. This step is easy to miss, and skipping it produces a `403` even though the credentials are valid.

Pass the key with `--sa-cred-file`:

```bash
# list the tab names
table-transform --list gsheets:1AbC --sa-cred-file service-account.json

# google sheets -> file
table-transform gsheets:1AbC/Roster export.csv --sa-cred-file service-account.json
table-transform gsheets:1AbC#0 export.json --sa-cred-file service-account.json

# every tab -> one xlsx workbook
table-transform gsheets:1AbC export.xlsx --sa-cred-file service-account.json

# file -> google sheets (the tab is created when missing)
table-transform source.csv gsheets:1AbC/Roster --sa-cred-file service-account.json
```

Column selection and trimming options work the same as for files.

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

convert_file({
  source_file: 'res/roster.md',
  dest_file: 'res/roster-selected.csv',
  fields: ['No', 'SID', 'Name'],
})
```

### Pick selected columns

Match columns by header name. When neither option is set, all columns are kept. Include (`fields`) and exclude (`exclude_fields`) cannot be used together in a single call (logically conflicting).

Use `fields` / `exclude_fields` on `read_file` or `convert_file` (see example above). Use `filter_fields` when you already have `rows` in memory:

```typescript
import { read_file, filter_fields } from 'table-transform'

let rows_1 = read_file({
  file: 'res/roster.md',
  fields: ['No', 'SID', 'Name'],
})[0].rows

let rows_2 = filter_fields({
  rows: rows_1,
  exclude_fields: ['No'],
})
// rows_2[0] → ['SID', 'Name']
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

### List sheet/table names

```typescript
import { list_sheet_names } from 'table-transform'

let names = list_sheet_names('res/multi-table.xlsx')
// ['Users', 'Posts']
```

### Read and write Google Sheets

Google Sheets needs network access, so it has its own async functions alongside the synchronous file ones. They take the same options, so a spreadsheet reads and writes like any other table.

```typescript
import {
  is_gsheets_ref,
  list_gsheets_names,
  read_file,
  read_gsheets_file,
  write_file,
  write_gsheets_file,
} from 'table-transform'

let sa_cred_file = 'service-account.json'

let names = await list_gsheets_names({ file: 'gsheets:1AbC', sa_cred_file })
// ['Users', 'Posts']

// spreadsheet -> csv
let sheets = await read_gsheets_file({
  file: 'gsheets:1AbC/Roster',
  fields: ['No', 'SID'],
  sa_cred_file,
})
write_file({ file: 'res/roster.csv', sheets })

// csv -> spreadsheet
await write_gsheets_file({
  file: 'gsheets:1AbC/Roster',
  sheets: read_file({ file: 'res/roster.csv' }),
  sa_cred_file,
})

// use is_gsheets_ref() when the path comes from user input
let rows = is_gsheets_ref(file)
  ? await read_gsheets_file({ file, sa_cred_file })
  : read_file({ file })
```

The synchronous `read_file` / `write_file` / `convert_file` / `list_sheet_names` handle local files only; passing them a `gsheets:` reference throws and points at the matching Google Sheets function.

A write replaces the whole tab: it is resized to the extent of the new rows, then every cell is written. A tab named in the reference is created if it does not exist yet, while `gsheets:<id>#<index>` requires an existing tab.

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
// [{ No: '1', Name: 'Alice Wong', ... }, ...]
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
  /** only include specified fields, exclude all other fields */
  fields?: string[]
  /** exclude specified fields, include all other fields */
  exclude_fields?: string[]
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

export function list_sheet_names(file: string): string[]

/* Google Sheets. These need network access, so they are async and sit beside
 * the synchronous file functions rather than inside them. Use `is_gsheets_ref()`
 * to tell which of the two a given path needs. */

export type GoogleSheetsOptions = {
  /** service account credential json file */
  sa_cred_file?: string
}

export type GoogleSheetsRef = {
  spreadsheet_id: string
  /** set when the tab is selected by name */
  sheet_name?: string
  /** set when the tab is selected by zero-based index */
  sheet_index?: number
}

export function is_gsheets_ref(file: string): boolean

/** returns null when the given string is an ordinary file path */
export function parse_gsheets_ref(file: string): GoogleSheetsRef | null

/** throws when the credential file is missing or not set */
export function resolve_sa_cred_file(options?: GoogleSheetsOptions): string

/** quote a tab name for A1 notation, e.g. `Sales 'Q1'` -> `'Sales ''Q1'''` */
export function encode_gsheets_sheet_name(name: string): string

export function list_gsheets_names(
  args: {
    file: string
  } & GoogleSheetsOptions,
): Promise<string[]>

export function read_gsheets_file(
  args: {
    file: string
  } & ExtraReadFileOptions &
    GoogleSheetsOptions,
): Promise<SheetData<CellValue>[]>

export function write_gsheets_file(
  args: {
    file: string
    sheets: SheetData<CellValue>[]
  } & GoogleSheetsOptions,
): Promise<void>

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

export function filter_fields<T extends CellValue>(args: {
  rows: T[][]
  fields?: string[]
  exclude_fields?: string[]
}): T[][]

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
