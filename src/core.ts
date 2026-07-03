import { readFile, writeFile, utils } from '@e965/xlsx'
import { to_csv, from_csv } from '@beenotung/tslib/csv'
import { readFileSync, writeFileSync } from 'fs'
import { basename, dirname, extname, join } from 'path'

export type SheetData<T extends CellValue> = {
  name: string
  rows: T[][]
}

export type CellValue = string | number | boolean | Date | null

type ExtraReadFileOptions = {
  /** for csv/txt, default auto detect ',' or '\t' */
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
  args: {
    file: string
  } & ExtraReadFileOptions,
): SheetData<CellValue>[] {
  let file = args.file
  let ext = extname(file)
  switch (ext) {
    case '.xlsx': {
      return read_xlsx_file(args)
    }
    case '.txt':
    case '.csv': {
      return [read_csv_file(args)]
    }
    case '.tsv': {
      return [read_csv_file({ ...args, separator: args.separator || '\t' })]
    }
    case '.markdown':
    case '.md': {
      return read_md_file(args)
    }
    case '.json': {
      return [read_json_file(args)]
    }
    default:
      throw new Error(`Unsupported file extension: ${ext}`)
  }
}

function infer_sheet_name(file: string) {
  let filename = basename(file)
  let ext = extname(filename)
  return filename.slice(0, -ext.length).trim()
}

export type ShowNameMode = 'auto' | 'always' | 'never'

export function write_file(args: {
  file: string
  sheets: SheetData<CellValue>[]
  /**
   * only used when file is /dev/stdout
   * @default 'auto'
   */
  show_name?: ShowNameMode
  /**
   * for txt files
   */
  separator?: string
  /**
   * for json files
   * @default 'object'
   */
  json_format?: 'object' | 'array'
}) {
  let { file, sheets, json_format } = args
  let ext = extname(file)
  let write: (file: string, rows: CellValue[][]) => void
  switch (ext) {
    case '.xlsx': {
      write_xlsx_file({ file, sheets })
      return
    }
    case '.txt': {
      write = (file, rows) =>
        write_csv_file({ file, rows, separator: args.separator || ' | ' })
      break
    }
    case '.csv': {
      write = (file, rows) => write_csv_file({ file, rows, separator: ',' })
      break
    }
    case '.tsv': {
      write = (file, rows) => write_csv_file({ file, rows, separator: '\t' })
      break
    }
    case '.markdown':
    case '.md': {
      write = (file, rows) => write_md_file({ file, rows })
      break
    }
    case '.json': {
      write = (file, rows) =>
        write_json_file({ file, rows, format: json_format })
      break
    }
    default: {
      throw new Error(`Unsupported file extension: ${ext}`)
    }
  }
  if (file.startsWith('/dev/stdout')) {
    file = '/dev/stdout'
  }
  let files = infer_dest_files({ file, sheets })
  let show_name = resolve_show_name({
    file,
    show_name: args.show_name || 'auto',
    files,
  })
  for (let i = 0; i < files.length; i++) {
    if (show_name) {
      console.log()
      console.log(`Table: ${sheets[i].name}`)
      console.log()
    }
    write(files[i], sheets[i].rows)
  }
}

function resolve_show_name(args: {
  file: string
  show_name: ShowNameMode
  files: string[]
}): boolean {
  if (args.file !== '/dev/stdout') {
    return false
  }
  if (args.show_name === 'auto') {
    return args.files.length > 1
  }
  if (args.show_name === 'always') {
    return true
  }
  if (args.show_name === 'never') {
    return false
  }
  throw new Error(`Invalid show name mode: ${args.show_name}`)
}

function infer_dest_files(args: {
  file: string
  sheets: SheetData<CellValue>[]
}): string[] {
  let { file, sheets } = args
  if (sheets.length === 1) {
    return [file]
  }
  if (file === '/dev/stdout') {
    return new Array(sheets.length).fill('/dev/stdout')
  }
  let dir = dirname(file)
  let name = infer_sheet_name(file)
  let ext = extname(file)
  return sheets.map(sheet =>
    join(dir, `${sanitize_name(name)}-${sanitize_name(sheet.name)}${ext}`),
  )
}

function sanitize_name(name: string): string {
  return name
    .replaceAll('/', '_')
    .replaceAll(':', '_')
    .replaceAll('\\', '_')
    .trim()
}

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
) {
  let {
    source_file,
    dest_file,
    json_format,
    input_separator,
    output_separator,
    separator,
    ...rest
  } = args
  input_separator ||= separator
  output_separator ||= separator
  let sheets = read_file({
    ...rest,
    file: source_file,
    separator: input_separator,
  })
  write_file({
    ...rest,
    file: dest_file,
    sheets,
    json_format,
    separator: output_separator,
  })
}

export type SheetInfo = {
  name: string
  range?: string
}

export function read_xlsx_file(
  args: {
    file: string
    sheets?: string[] | SheetInfo[]
  } & ExtraReadFileOptions,
): SheetData<CellValue>[] {
  let file = args.file
  let workbook = readFile(file)
  let sheets: SheetInfo[] = (args.sheets || workbook.SheetNames).map(item => {
    return typeof item === 'string' ? { name: item } : item
  })
  return sheets.map(info => {
    let sheet = workbook.Sheets[info.name]
    let range = utils.decode_range(info.range || sheet['!ref'] || 'A1:A1')
    let rows: CellValue[][] = []
    for (let r = range.s.r; r <= range.e.r; r++) {
      let cols: CellValue[] = []
      for (let c = range.s.c; c <= range.e.c; c++) {
        let cell = sheet[utils.encode_cell({ r, c })]
        if (cell) {
          cols.push(cell.v)
        } else {
          cols.push(null)
        }
      }
      rows.push(cols)
    }
    rows = trim_rows(rows, args)
    return { name: info.name, rows }
  })
}

function trim_rows<T extends CellValue>(
  rows: T[][],
  args: ExtraReadFileOptions,
): T[][] {
  if (args.trim_string ?? true) {
    rows = trim_string(rows)
  }
  if ((args.trim_rows ?? true) || (args.trim_cols ?? true)) {
    rows = trim_cell(rows, args)
  }
  return rows
}

function trim_string<T extends CellValue>(rows: T[][]): T[][] {
  return rows.map(cols =>
    cols.map(value =>
      typeof value === 'string' ? (value.trim() as T) : value,
    ),
  )
}

function trim_cell<T extends CellValue>(
  rows: T[][],
  args: { trim_rows?: boolean; trim_cols?: boolean },
): T[][] {
  if (rows.length === 0) {
    return rows
  }
  let trim_rows = args.trim_rows ?? true
  let trim_cols = args.trim_cols ?? true
  let start_row = 0
  let end_row = rows.length - 1
  let start_col = 0
  let end_col = rows[0].length - 1

  for (; end_row >= start_row && end_col >= start_col; ) {
    // trim tailing row
    let has_value = false
    for (let c = start_col; trim_rows && c <= end_col; c++) {
      let value = rows[end_row][c]
      if (value !== null && value !== '') {
        has_value = true
        break
      }
    }
    if (trim_rows && !has_value) {
      end_row--
      continue
    }

    // trim tailing col
    has_value = false
    for (let r = start_row; trim_cols && r <= end_row; r++) {
      let value = rows[r][end_col]
      if (value !== null && value !== '') {
        has_value = true
        break
      }
    }
    if (trim_cols && !has_value) {
      end_col--
      continue
    }

    // trim leading row
    has_value = false
    for (let c = start_col; trim_rows && c <= end_col; c++) {
      let value = rows[start_row][c]
      if (value !== null && value !== '') {
        has_value = true
        break
      }
    }
    if (trim_rows && !has_value) {
      start_row++
      continue
    }

    // trim leading col
    has_value = false
    for (let r = start_row; trim_cols && r <= end_row; r++) {
      let value = rows[r][start_col]
      if (value !== null && value !== '') {
        has_value = true
        break
      }
    }
    if (trim_cols && !has_value) {
      start_col++
      continue
    }

    break
  }

  if (end_row < start_row || end_col < start_col) {
    return []
  }

  return rows
    .slice(start_row, end_row + 1)
    .map(row => row.slice(start_col, end_col + 1))
}

export function read_csv_file(
  args: {
    file: string
  } & ExtraReadFileOptions,
): SheetData<string> & { separator: string } {
  let file = args.file
  let text = readFileSync(file, 'utf-8')

  let separator = args.separator
  if (!separator && file.endsWith('.tsv')) {
    separator = '\t'
  }
  let rows: string[][]
  if (separator) {
    rows = from_csv(text, separator)
  } else {
    let result = infer_csv_separator(text)
    rows = result.rows
    separator = result.separator
  }
  rows = trim_rows(rows, args)
  let name = infer_sheet_name(file)
  return { name, rows, separator }
}

export function infer_csv_separator(text: string): {
  separator: ',' | '\t'
  rows: string[][]
} {
  let comma = from_csv(text, ',')
  let tab = from_csv(text, '\t')

  let comma_count = 0
  for (let row of comma) {
    comma_count += row.length
  }

  let tab_count = 0
  for (let row of tab) {
    tab_count += row.length
  }

  return tab_count > comma_count
    ? { separator: '\t', rows: tab }
    : { separator: ',', rows: comma }
}

export function read_md_file(
  args: {
    file: string
  } & ExtraReadFileOptions,
): SheetData<string>[] {
  let file = args.file
  let name = infer_sheet_name(file)
  let text = readFileSync(file, 'utf-8')
  let lines = text.split('\n')
  let sheets: SheetData<string>[] = []
  let offset = 0
  while (offset < lines.length) {
    let start_index = find_index({
      lines,
      offset,
      predicate: line => line.startsWith('|'),
    })
    if (start_index === -1) {
      break
    }
    let end_index = find_index({
      lines,
      offset: start_index + 1,
      predicate: line => !line.startsWith('|'),
    })
    if (end_index === -1) {
      end_index = lines.length
    }
    let rows: string[][] = []
    for (let i = start_index; i < end_index; i++) {
      rows.push(parse_md_line(lines[i]))
    }
    rows.splice(1, 1)
    rows = trim_rows(rows, args)
    sheets.push({ name: name + `-${sheets.length + 1}`, rows })
    offset = end_index
  }
  if (sheets.length === 1) {
    sheets[0].name = name
  }
  return sheets
}

function find_index(args: {
  lines: string[]
  offset: number
  predicate: (line: string) => boolean
}): number {
  let { lines, offset, predicate } = args
  for (let i = offset; i < lines.length; i++) {
    if (predicate(lines[i])) {
      return i
    }
  }
  return -1
}

function parse_md_line(line: string) {
  let cells: string[] = []
  let buffer = ''
  for (let i = 1; i < line.length; i++) {
    switch (line[i]) {
      case '|': {
        cells.push(
          buffer
            .replaceAll('<br>', '\n')
            .replaceAll('<br/>', '\n')
            .replaceAll('<br />', '\n')
            .trim(),
        )
        buffer = ''
        break
      }
      case '\\': {
        i++
        buffer += line[i] || ''
        break
      }
      default: {
        buffer += line[i]
      }
    }
  }
  return cells
}

export function write_csv_file(args: {
  file: string
  rows: CellValue[][]
  separator?: string
}) {
  let file = args.file
  let separator = args.separator
  if (!separator && file.endsWith('.tsv')) {
    separator = '\t'
  }
  if (!separator && file.endsWith('.txt')) {
    separator = ' | '
  }
  if (!separator) {
    separator = ','
  }
  let newline = file.endsWith('.txt') ? '\\n' : '\n'
  let text = to_csv(
    args.rows.map(row => row.map(col => value_to_string(col, newline))),
    separator,
  )
  writeFileSync(file, text, 'utf-8')
}

function value_to_string(value: CellValue, line_break: string) {
  if (value === null) {
    return ''
  }
  if (typeof value === 'string') {
    return line_break === '\n' ? value : value.replaceAll('\n', line_break)
  }
  if (value instanceof Date) {
    return value.toISOString()
  }
  return String(value)
}

export function write_xlsx_file(args: {
  file: string
  sheets: SheetData<CellValue>[]
}) {
  let file = args.file
  let workbook = utils.book_new()
  for (let { name, rows } of args.sheets) {
    let sheet = utils.aoa_to_sheet(rows)
    utils.book_append_sheet(workbook, sheet, name)
  }
  writeFile(workbook, file)
}

export function write_md_file(args: { file: string; rows: CellValue[][] }) {
  let file = args.file
  let rows = args.rows.map(cols =>
    cols.map(col => value_to_string(col, '<br>')),
  )

  let lengths = rows[0].map(col => col.length)
  for (let r = 1; r < rows.length; r++) {
    let cols = rows[r]
    for (let c = 0; c < cols.length; c++) {
      let length = count_char_width(cols[c])
      if (length > lengths[c]) {
        lengths[c] = length
      }
    }
  }

  for (let i = 0; i < lengths.length; i++) {
    if (lengths[i] < 3) {
      lengths[i] = 3
    }
  }

  let text = ''
  for (let r = 0; r < rows.length; r++) {
    let row = rows[r]
    text += '|'
    for (let c = 0; c < row.length; c++) {
      let cell_length = lengths[c]
      let value = row[c]
      let value_length = count_char_width(value)
      if (value_length < cell_length) {
        value += ' '.repeat(cell_length - value_length)
      }
      text += ` ${value} |`
    }
    text += '\n'
    if (r === 0) {
      text += '|'
      for (let c = 0; c < row.length; c++) {
        let length = lengths[c]
        text += ` ${'-'.repeat(length)} |`
      }
      text += '\n'
    }
  }
  writeFileSync(file, text, 'utf-8')
}

function count_char_width(text: string) {
  let width = 0
  for (let char of text) {
    if (/[一-鿕|㐀-䷿|〱-㏿|가-힣|、-〠|！-｠|￠-￦|띀-렝]/.test(char)) {
      width += 2
    } else {
      width += 1
    }
  }
  return width
}

export function read_json_file(
  args: {
    file: string
    format?: 'array' | 'object'
  } & ExtraReadFileOptions,
): SheetData<CellValue> {
  let { file, format } = args
  let name = infer_sheet_name(file)
  let text = readFileSync(file, 'utf-8')
  let json = JSON.parse(text)
  switch (format) {
    case 'array': {
      if (!Array.isArray(json)) {
        throw new Error('Invalid JSON file, expected array')
      }
      for (let item of json) {
        if (!Array.isArray(item)) {
          throw new Error('Invalid JSON file, expected array of arrays')
        }
      }
      let rows = trim_rows(json, args)
      return { name, rows }
    }
    case 'object': {
      if (is_object(json)) {
        json = [json]
      }
      if (!Array.isArray(json)) {
        throw new Error('Invalid JSON file, expected array of objects')
      }
      for (let item of json) {
        if (!is_object(item)) {
          throw new Error('Invalid JSON file, expected array of objects')
        }
      }
      let rows = objects_to_rows<CellValue>(json)
      rows = trim_rows(rows, args)
      return { name, rows }
    }
    default: {
      // infer format from data

      if (is_object(json)) {
        json = [json]
      }

      if (!Array.isArray(json)) {
        throw new Error('Invalid JSON file, expected array')
      }

      for (let item of json) {
        if (Array.isArray(item)) {
          return read_json_file({ ...args, file, format: 'array' })
        }
        if (is_object(item)) {
          return read_json_file({ ...args, file, format: 'object' })
        }
        throw new Error(
          'Invalid JSON file, expected array of objects or arrays',
        )
      }

      return { name, rows: [] }
    }
  }
}

export function write_json_file(args: {
  file: string
  rows: CellValue[][]
  /** @default 'object' */
  format?: 'array' | 'object'
}) {
  let { file, rows } = args
  let format = args.format || 'object'
  let values = format === 'object' ? rows_to_objects(rows) : rows
  let text = '['
  for (let i = 0; i < values.length; i++) {
    if (i !== 0) {
      text += ','
    }
    text += '\n  ' + JSON.stringify(values[i])
  }
  text += '\n]\n'
  writeFileSync(file, text, 'utf-8')
}

function is_object(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

export function rows_to_objects<T>(rows: T[][]): Record<string, T>[] {
  if (rows.length === 0) {
    return []
  }
  let objects: Record<string, T>[] = []
  let headers = rows[0]
  for (let i = 1; i < rows.length; i++) {
    let row = rows[i]
    let object: Record<string, T> = {}
    for (let j = 0; j < headers.length; j++) {
      object[headers[j] as string] = row[j]
    }
    objects.push(object)
  }
  return objects
}

export function objects_to_rows<T>(
  objects: Record<string, T>[],
): (T | string | null)[][] {
  if (objects.length === 0) {
    return []
  }
  let header_set = new Set<string>()
  for (let object of objects) {
    for (let header of Object.keys(object)) {
      header_set.add(header)
    }
  }
  let headers = Array.from(header_set)
  let rows: (T | string | null)[][] = [headers]
  for (let object of objects) {
    let row: (T | string | null)[] = []
    for (let header of headers) {
      if (header in object) {
        row.push(object[header as keyof typeof object])
      } else {
        row.push(null)
      }
    }
    rows.push(row)
  }
  return rows
}
