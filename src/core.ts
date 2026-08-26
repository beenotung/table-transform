import { readFile, writeFile, utils } from '@e965/xlsx'
import { to_csv, from_csv } from '@beenotung/tslib/csv'
import { existsSync, readFileSync, writeFileSync } from 'fs'
import { basename, dirname, extname, join } from 'path'
// type-only, the module itself is loaded on demand, see create_gsheets_client()
import type { sheets_v4 } from '@googleapis/sheets'

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
  args: {
    file: string
  } & ExtraReadFileOptions,
): SheetData<CellValue>[] {
  let file = args.file
  require_local_file(file, 'read_file_async')
  let ext = extname(file)
  let sheets: SheetData<CellValue>[]
  switch (ext) {
    case '.xlsx': {
      sheets = read_xlsx_file(args)
      break
    }
    case '.txt': {
      sheets = read_txt_file(args)
      break
    }
    case '.csv': {
      sheets = [read_csv_file(args)]
      break
    }
    case '.tsv': {
      sheets = [read_csv_file({ ...args, separator: args.separator || '\t' })]
      break
    }
    case '.markdown':
    case '.md': {
      sheets = read_md_file(args)
      break
    }
    case '.json': {
      sheets = [read_json_file(args)]
      break
    }
    default:
      throw new Error(`Unsupported file extension: ${ext}`)
  }
  return sheets.map(sheet => ({
    ...sheet,
    rows: filter_fields({
      rows: sheet.rows,
      fields: args.fields,
      exclude_fields: args.exclude_fields,
    }),
  }))
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
  require_local_file(file, 'write_file_async')
  let ext = extname(file)
  let write: (file: string, rows: CellValue[][]) => void
  switch (ext) {
    case '.xlsx': {
      write_xlsx_file({ file, sheets })
      return
    }
    case '.txt': {
      write = (file, rows) => write_txt_file({ file, rows })
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
  require_local_file(source_file, 'convert_file_async')
  require_local_file(dest_file, 'convert_file_async')
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

/** List sheet/table names without converting cell data. */
export function list_sheet_names(file: string): string[] {
  require_local_file(file, 'list_sheet_names_async')
  let ext = extname(file)
  switch (ext) {
    case '.xlsx': {
      // SheetJS: parse sheet names only, skip cell data
      let workbook = readFile(file, { bookSheets: true })
      return workbook.SheetNames
    }
    case '.csv':
    case '.tsv':
    case '.json': {
      // single sheet named after the file basename
      return [infer_sheet_name(file)]
    }
    default: {
      // md/txt (and any future multi-table formats via read_file)
      return read_file({ file }).map(sheet => sheet.name)
    }
  }
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

export function read_txt_file(args: { file: string }): SheetData<string>[] {
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
      predicate: line => line.startsWith('+-'),
    })
    if (start_index === -1) {
      break
    }
    let end_index = find_index({
      lines,
      offset: start_index + 3,
      predicate: line => line.startsWith('+-'),
    })
    if (end_index === -1) {
      end_index = lines.length
    }
    let rows: string[][] = []
    // parse headers row
    rows.push(parse_txt_line(lines[start_index + 1]))
    // parse value rows
    for (let i = start_index + 3; i < end_index; i++) {
      rows.push(parse_txt_line(lines[i]))
    }
    sheets.push({ name: name + `-${sheets.length + 1}`, rows })
    offset = end_index + 1
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

function parse_txt_line(line: string) {
  let cells: string[] = []
  let buffer = ''
  for (let i = 1; i < line.length; i++) {
    switch (line[i]) {
      case '|': {
        cells.push(
          buffer
            .replaceAll('\\t', '\t')
            .replaceAll('\\r', '\r')
            .replaceAll('\\n', '\n')
            .replaceAll('\\\\', '\\')
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

function count_lengths(rows: string[][]): number[] {
  if (rows.length === 0) {
    return []
  }

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
  return lengths
}

export function write_md_file(args: { file: string; rows: CellValue[][] }) {
  if (args.rows.length === 0) {
    return
  }

  let rows = args.rows.map(cols =>
    cols.map(col => value_to_string(col, '<br>')),
  )

  let lengths = count_lengths(rows)

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
  writeFileSync(args.file, text, 'utf-8')
}

type ColType = 'null' | 'int' | 'real' | 'text'

function infer_col_types(rows: CellValue[][]): ColType[] {
  if (rows.length === 0) {
    return []
  }

  let types: ColType[] = []
  for (let r = 1; r < rows.length; r++) {
    let row = rows[r]
    for (let c = 0; c < row.length; c++) {
      let value = row[c]
      let type = to_col_type(value)
      if (types[c] === type) {
        continue
      }
      if (!types[c] || types[c] === 'null') {
        types[c] = type
        continue
      }
      if (types[c] === 'int' && type === 'real') {
        types[c] = 'real'
        continue
      }
      if (type === 'text') {
        types[c] = 'text'
        continue
      }
    }
  }
  return types
}

function to_col_type(value: CellValue): ColType {
  if (value === null || value === '') {
    return 'null'
  }
  if (typeof value === 'string') {
    let num = +value
    let str = String(num)
    if (str === value || str + '.0' === value) {
      value = num
    }
  }
  if (typeof value === 'number') {
    if (Number.isInteger(value)) {
      return 'int'
    }
    if (Number.isFinite(value)) {
      return 'real'
    }
    return 'int'
  }
  return 'text'
}

export function write_txt_file(args: { file: string; rows: CellValue[][] }) {
  let value_rows = args.rows
  if (value_rows.length === 0) {
    return
  }

  let str_rows = value_rows.map(cols =>
    cols.map(col =>
      value_to_string(col, '\n')
        .replaceAll('\\', '\\\\')
        .replaceAll('\t', '\\t')
        .replaceAll('\r', '\\r')
        .replaceAll('\n', '\\n'),
    ),
  )

  let lengths = count_lengths(str_rows)
  let types = infer_col_types(value_rows)

  let text = ''

  let border =
    '+' + lengths.map(length => '-'.repeat(length + 2)).join('+') + '+'

  text += border + '\n'

  for (let r = 0; r < str_rows.length; r++) {
    let row = str_rows[r]
    text += '|'
    for (let c = 0; c < row.length; c++) {
      let type = types[c]
      let align =
        r === 0
          ? 'center'
          : type === 'int' || type === 'real'
            ? 'left'
            : 'right'
      let cell_length = lengths[c]
      let str = row[c]
      let str_length = count_char_width(str)
      let padding = cell_length - str_length
      if (padding > 0) {
        if (align === 'center') {
          let left_padding = Math.floor(padding / 2)
          let right_padding = padding - left_padding
          str = ' '.repeat(left_padding) + str + ' '.repeat(right_padding)
        } else if (align === 'right') {
          str = str + ' '.repeat(padding)
        } else if (align === 'left') {
          str = ' '.repeat(padding) + str
        }
      }
      text += ` ${str} |`
    }
    text += '\n'
    if (r === 0) {
      text += border + '\n'
    }
  }

  text += border + '\n'

  writeFileSync(args.file, text, 'utf-8')
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

export function filter_fields<T extends CellValue>(args: {
  rows: T[][]
  fields?: string[]
  exclude_fields?: string[]
}): T[][] {
  let { rows, fields, exclude_fields: excluded } = args
  if (rows.length === 0) {
    return rows
  }
  if (!fields?.length && !excluded?.length) {
    return rows
  }
  if (fields?.length && excluded?.length) {
    throw new Error('Cannot filter by both including and excluding fields')
  }
  let headers = rows[0].map(header => String(header))
  let indices: number[] = []
  if (fields?.length) {
    for (let name of fields) {
      let index = headers.indexOf(name)
      if (index === -1) {
        console.warn(
          `Skip field: ${JSON.stringify(name)}, not found in headers`,
        )
        continue
      }
      indices.push(index)
    }
  } else {
    for (let i = 0; i < headers.length; i++) {
      if (!excluded?.includes(headers[i])) {
        indices.push(i)
      }
    }
  }
  return rows.map(row => indices.map(index => row[index]))
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

/*
 * Referenced by a `gsheets:` pseudo-path instead of a file path:
 *
 *   gsheets:<spreadsheet-id>                 all tabs
 *   gsheets:<spreadsheet-id>/<sheet-name>    one tab, by name
 *   gsheets:<spreadsheet-id>#<sheet-index>   one tab, by zero-based index
 *
 * Google Sheets requires network access, so it is only reachable through the
 * async variants (`read_file_async`, `write_file_async`, `convert_file_async`,
 * `list_sheet_names_async`). Those delegate to the original synchronous
 * functions for ordinary file paths, so existing behaviour is unchanged.
 */

export type GoogleSheetsRef = {
  spreadsheet_id: string
  /** set when the tab is selected by name */
  sheet_name?: string
  /** set when the tab is selected by zero-based index */
  sheet_index?: number
}

export type GoogleSheetsOptions = {
  /** service account credential json file */
  sa_cred_file?: string
}

export function is_gsheets_ref(file: string): boolean {
  return file.startsWith('gsheets:')
}

function require_local_file(file: string, async_fn: string): void {
  if (is_gsheets_ref(file)) {
    throw new Error(
      `Google Sheets requires the async API, use ${async_fn}() instead`,
    )
  }
}

/**
 * Parse a `gsheets:` reference, or return null for an ordinary file path. The
 * id ends at the first "/" or "#", so a tab name may contain "#".
 */
export function parse_gsheets_ref(file: string): GoogleSheetsRef | null {
  if (!is_gsheets_ref(file)) {
    return null
  }
  let [, id, separator, selector] = /^gsheets:([^/#]*)(?:([/#])(.*))?$/s.exec(
    file,
  )!
  let spreadsheet_id = id.trim()
  if (!spreadsheet_id) {
    throw new Error('Missing Google Sheets spreadsheet id')
  }
  if (!separator) {
    return { spreadsheet_id }
  }
  selector = selector.trim()
  if (separator === '/') {
    if (!selector) {
      throw new Error('Missing Google Sheets sheet name after "/"')
    }
    return { spreadsheet_id, sheet_name: selector }
  }
  let sheet_index = +selector
  if (!/^\d+$/.test(selector) || !Number.isSafeInteger(sheet_index)) {
    throw new Error(
      'Invalid Google Sheets sheet index after "#", expected a zero-based non-negative integer',
    )
  }
  return { spreadsheet_id, sheet_index }
}

function require_gsheets_ref(file: string): GoogleSheetsRef {
  let ref = parse_gsheets_ref(file)
  if (!ref) {
    throw new Error(`Not a Google Sheets reference: ${file}`)
  }
  return ref
}

/**
 * Resolve the service account credential file, checking it here so a wrong
 * path is reported before any network call, where it would be a bare ENOENT.
 */
export function resolve_sa_cred_file(
  options: GoogleSheetsOptions = {},
): string {
  let key_file = options.sa_cred_file
  if (!key_file) {
    throw new Error(
      'Missing Google service account credential file, use --sa-cred-file <file>',
    )
  }
  if (!existsSync(key_file)) {
    throw new Error(
      `Google service account credential file not found: ${key_file}`,
    )
  }
  return key_file
}

type GoogleSheetsClient = sheets_v4.Sheets

/** a tab, by its stable numeric id (needed to resize it) and its name */
type GoogleSheetsGrid = { id: number; name: string }

function to_gsheets_grid(properties: any): GoogleSheetsGrid {
  let name = String(properties?.title ?? '')
  let id = properties?.sheetId
  // not defaulted, 0 is a real tab id and a resize would hit the wrong tab
  if (typeof id !== 'number' || !Number.isInteger(id)) {
    throw new Error(
      `Missing Google Sheets tab id in the API response for tab: ${name}`,
    )
  }
  return { id, name }
}

/** unwrap a Sheets API response, naming the operation in a failure */
async function gsheets_call<T>(
  operation: string,
  request: () => Promise<{ data: T }>,
): Promise<T> {
  try {
    return (await request()).data
  } catch (error: any) {
    let status =
      typeof error?.status === 'number' ? ` (HTTP ${error.status})` : ''
    throw new Error(
      `Failed to ${operation} in Google Sheets${status}: ${error?.message || error}`,
      { cause: error },
    )
  }
}

/** Authenticate against the spreadsheet and list its tabs. */
async function open_gsheets(args: {
  ref: GoogleSheetsRef
  options: GoogleSheetsOptions
  readonly: boolean
}) {
  let spreadsheet_id = args.ref.spreadsheet_id
  let key_file = resolve_sa_cred_file(args.options)
  // loaded on demand so non-gsheets usage never pays for it
  let { sheets, auth } = await import('@googleapis/sheets')
  let scope = 'https://www.googleapis.com/auth/spreadsheets'
  let client: GoogleSheetsClient = sheets({
    version: 'v4',
    auth: new auth.GoogleAuth({
      keyFile: key_file,
      scopes: [args.readonly ? scope + '.readonly' : scope],
    }),
  })
  let data = await gsheets_call('list the tabs', () =>
    client.spreadsheets.get({
      spreadsheetId: spreadsheet_id,
      fields: 'sheets.properties(sheetId,title)',
    }),
  )
  let grids = (data.sheets || []).map(sheet =>
    to_gsheets_grid(sheet?.properties),
  )
  return { spreadsheet_id, client, grids }
}

/** the tabs the reference selects, or every tab when it has no selector */
function resolve_gsheets_grids(
  grids: GoogleSheetsGrid[],
  ref: GoogleSheetsRef,
): GoogleSheetsGrid[] {
  if (ref.sheet_index !== undefined) {
    let grid = grids[ref.sheet_index]
    if (!grid) {
      throw new Error(
        `Google Sheets tab index out of range: ${ref.sheet_index}, the spreadsheet has ${grids.length} tab(s)`,
      )
    }
    return [grid]
  }
  if (ref.sheet_name !== undefined) {
    let grid = grids.find(grid => grid.name === ref.sheet_name)
    if (!grid) {
      throw new Error(`Google Sheets tab not found: ${ref.sheet_name}`)
    }
    return [grid]
  }
  return grids
}

/** quote a tab name for A1 notation, e.g. `Sales 'Q1'` -> `'Sales ''Q1'''` */
export function encode_gsheets_sheet_name(name: string): string {
  return `'${name.replaceAll("'", "''")}'`
}

/** pad ragged rows into a rectangle in place, and report the width */
function squarize_rows<T>(rows: T[][], filler: T): number {
  let width = rows.reduce((width, row) => Math.max(width, row.length), 0)
  for (let row of rows) {
    while (row.length < width) {
      row.push(filler)
    }
  }
  return width
}

function from_gsheets_value(value: unknown): CellValue {
  let type = typeof value
  if (type === 'string' || type === 'number' || type === 'boolean') {
    return value as CellValue
  }
  return value === null || value === undefined ? null : String(value)
}

function to_gsheets_value(value: CellValue): string | number | boolean {
  if (value === null) {
    return ''
  }
  return value instanceof Date ? value.toISOString() : value
}

/** List the tab names of a Google Sheets spreadsheet. */
export async function list_gsheets_names(
  args: {
    file: string
  } & GoogleSheetsOptions,
): Promise<string[]> {
  let ref = require_gsheets_ref(args.file)
  let { grids } = await open_gsheets({ ref, options: args, readonly: true })
  return resolve_gsheets_grids(grids, ref).map(grid => grid.name)
}

/** Read one or more tabs of a Google Sheets spreadsheet. */
export async function read_gsheets_file(
  args: {
    file: string
  } & ExtraReadFileOptions &
    GoogleSheetsOptions,
): Promise<SheetData<CellValue>[]> {
  let ref = require_gsheets_ref(args.file)
  let { spreadsheet_id, client, grids } = await open_gsheets({
    ref,
    options: args,
    readonly: true,
  })
  let sheets: SheetData<CellValue>[] = []
  // read tabs one at a time to bound memory and stay clear of rate limits
  for (let grid of resolve_gsheets_grids(grids, ref)) {
    let data = await gsheets_call(`read the tab: ${grid.name}`, () =>
      client.spreadsheets.values.get({
        spreadsheetId: spreadsheet_id,
        // an unqualified tab name covers the whole tab
        range: encode_gsheets_sheet_name(grid.name),
        majorDimension: 'ROWS',
        valueRenderOption: 'UNFORMATTED_VALUE',
        dateTimeRenderOption: 'FORMATTED_STRING',
      }),
    )
    let rows: CellValue[][] = (data.values || []).map(row =>
      (row || []).map(from_gsheets_value),
    )
    // padded so the trim and field helpers can index every column
    squarize_rows(rows, null)
    sheets.push({ name: grid.name, rows: trim_rows(rows, args) })
  }
  return sheets
}

async function ensure_gsheets_grid(args: {
  client: GoogleSheetsClient
  spreadsheet_id: string
  grids: GoogleSheetsGrid[]
  name: string
}): Promise<GoogleSheetsGrid> {
  let { client, spreadsheet_id, grids, name } = args
  let existing = grids.find(grid => grid.name === name)
  if (existing) {
    return existing
  }
  let data = await gsheets_call(`create the tab: ${name}`, () =>
    client.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheet_id,
      requestBody: {
        requests: [{ addSheet: { properties: { title: name } } }],
      },
    }),
  )
  let properties = data.replies?.[0]?.addSheet?.properties
  if (!properties) {
    throw new Error(`Failed to create Google Sheets tab: ${name}`)
  }
  let grid = to_gsheets_grid(properties)
  grids.push(grid)
  return grid
}

async function write_gsheets_grid(args: {
  client: GoogleSheetsClient
  spreadsheet_id: string
  grid: GoogleSheetsGrid
  rows: CellValue[][]
}): Promise<void> {
  let { client, spreadsheet_id, grid, rows } = args
  // short rows are padded, otherwise their trailing cells keep the old values
  let values = rows.map(row => row.map(to_gsheets_value))
  let width = squarize_rows<string | number | boolean>(values, '')
  // a tab always keeps at least one cell, so an empty table writes one blank
  if (width === 0) {
    values = [['']]
    width = 1
  }
  let height = values.length

  /*
   * Resized first, which both discards whatever lay outside the new extent and
   * grows the grid when the new rows do not fit, something values.update
   * cannot do: it rejects a range past the current grid limits.
   */
  let grid_properties = { rowCount: height, columnCount: width }
  await gsheets_call(`resize the tab: ${grid.name}`, () =>
    client.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheet_id,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: { sheetId: grid.id, gridProperties: grid_properties },
              fields: 'gridProperties.rowCount,gridProperties.columnCount',
            },
          },
        ],
      },
    }),
  )

  // A1 column letters are the same in Sheets and xlsx, so reuse the encoder
  let sheet_name = encode_gsheets_sheet_name(grid.name)
  let range = `${sheet_name}!A1:${utils.encode_col(width - 1)}${height}`
  await gsheets_call(`write the range: ${range}`, () =>
    client.spreadsheets.values.update({
      spreadsheetId: spreadsheet_id,
      range,
      valueInputOption: 'RAW',
      requestBody: { range, majorDimension: 'ROWS', values },
    }),
  )
}

/**
 * Write tables into a Google Sheets spreadsheet.
 *
 * `gsheets:<id>/<name>` and `gsheets:<id>#<index>` target a single tab, while
 * `gsheets:<id>` writes each table to its own tab named after the table.
 */
export async function write_gsheets_file(
  args: {
    file: string
    sheets: SheetData<CellValue>[]
  } & GoogleSheetsOptions,
): Promise<void> {
  let ref = require_gsheets_ref(args.file)
  let sheets = args.sheets
  if (sheets.length === 0) {
    return
  }
  let { spreadsheet_id, client, grids } = await open_gsheets({
    ref,
    options: args,
    readonly: false,
  })

  let names: string[]
  if (ref.sheet_index !== undefined) {
    // an index cannot name a new tab, so it must already exist
    names = resolve_gsheets_grids(grids, ref).map(grid => grid.name)
  } else if (ref.sheet_name !== undefined) {
    names = [ref.sheet_name]
  } else {
    names = sheets.map((sheet, i) => sheet.name || `Sheet${i + 1}`)
  }
  if (names.length < sheets.length) {
    throw new Error(
      `Cannot write ${sheets.length} tables into a single Google Sheets tab, use gsheets:${spreadsheet_id} to write each table to its own tab`,
    )
  }
  // two tables cannot share a tab, the second would overwrite the first
  let duplicated = names.find((name, i) => names.indexOf(name) !== i)
  if (duplicated !== undefined) {
    throw new Error(
      `Duplicated table name: ${duplicated}, each table needs its own Google Sheets tab`,
    )
  }

  for (let i = 0; i < sheets.length; i++) {
    let grid = await ensure_gsheets_grid({
      client,
      spreadsheet_id,
      grids,
      name: names[i],
    })
    await write_gsheets_grid({
      client,
      spreadsheet_id,
      grid,
      rows: sheets[i].rows,
    })
  }
}

/**
 * Async version of `list_sheet_names` that also accepts `gsheets:` references.
 */
export async function list_sheet_names_async(
  file: string,
  options: GoogleSheetsOptions = {},
): Promise<string[]> {
  if (is_gsheets_ref(file)) {
    return list_gsheets_names({ ...options, file })
  }
  return list_sheet_names(file)
}

/** Async version of `read_file` that also accepts `gsheets:` references. */
export function read_file_async(
  args: {
    /*
     * Listed first, otherwise a tab named like a file (`gsheets:<id>/data.csv`)
     * matches the text file overload below and is typed as string cells.
     */
    file: `gsheets:${string}`
  } & ExtraReadFileOptions &
    GoogleSheetsOptions,
): Promise<SheetData<CellValue>[]>
export function read_file_async(
  args: {
    file:
      | `${string}.md`
      | `${string}.markdown`
      | `${string}.csv`
      | `${string}.tsv`
      | `${string}.txt`
  } & ExtraReadFileOptions & GoogleSheetsOptions,
): Promise<SheetData<string>[]>
export function read_file_async(
  args: {
    file: `${string}.xlsx` | `${string}.json` | string
  } & ExtraReadFileOptions &
    GoogleSheetsOptions,
): Promise<SheetData<CellValue>[]>
export async function read_file_async(
  args: {
    file: string
  } & ExtraReadFileOptions &
    GoogleSheetsOptions,
): Promise<SheetData<CellValue>[]> {
  if (!is_gsheets_ref(args.file)) {
    return read_file(args)
  }
  let sheets = await read_gsheets_file(args)
  return sheets.map(sheet => ({
    ...sheet,
    rows: filter_fields({
      rows: sheet.rows,
      fields: args.fields,
      exclude_fields: args.exclude_fields,
    }),
  }))
}

/** Async version of `write_file` that also accepts `gsheets:` references. */
export async function write_file_async(
  args: {
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
  } & GoogleSheetsOptions,
): Promise<void> {
  if (is_gsheets_ref(args.file)) {
    return write_gsheets_file(args)
  }
  write_file(args)
}

/**
 * Async version of `convert_file`. Either side may be a `gsheets:` reference,
 * so this also converts a spreadsheet to a file and a file to a spreadsheet.
 */
export async function convert_file_async(
  args: {
    source_file: string
    dest_file: string
    /** for output json files, default 'object' */
    json_format?: 'object' | 'array'
    /** default same as separator */
    input_separator?: string
    /** default same as separator */
    output_separator?: string
  } & ExtraReadFileOptions &
    GoogleSheetsOptions,
): Promise<void> {
  let {
    source_file,
    dest_file,
    json_format,
    input_separator,
    output_separator,
    separator,
    sa_cred_file,
    ...rest
  } = args
  if (!is_gsheets_ref(source_file) && !is_gsheets_ref(dest_file)) {
    // sa_cred_file is dropped here, it means nothing to the local file path
    convert_file({
      ...rest,
      source_file,
      dest_file,
      json_format,
      input_separator,
      output_separator,
      separator,
    })
    return
  }
  input_separator ||= separator
  output_separator ||= separator
  let sheets = await read_file_async({
    ...rest,
    sa_cred_file,
    file: source_file,
    separator: input_separator,
  })
  await write_file_async({
    ...rest,
    sa_cred_file,
    file: dest_file,
    sheets,
    json_format,
    separator: output_separator,
  })
}
