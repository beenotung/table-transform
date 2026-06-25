import { readFile, writeFile, utils } from '@e965/xlsx'
import { to_csv, from_csv } from '@beenotung/tslib/csv'
import { readFileSync, writeFileSync } from 'fs'
import { basename, extname } from 'path'

export type SheetData<T extends CellValue> = {
  name: string
  rows: T[][]
}

export type CellValue = string | number | boolean | Date | null

export function read_file(args: {
  file:
    | `${string}.md`
    | `${string}.markdown`
    | `${string}.csv`
    | `${string}.tsv`
}): SheetData<string>[]
export function read_file(args: {
  file: `${string}.xlsx` | `${string}.json` | string
}): SheetData<CellValue>[]
export function read_file(args: { file: string }): SheetData<CellValue>[] {
  let file = args.file
  let ext = extname(file)
  switch (ext) {
    case '.xlsx': {
      return read_xlsx_file({ file })
    }
    case '.csv': {
      return [read_csv_file({ file })]
    }
    case '.tsv': {
      return [read_csv_file({ file, separator: '\t' })]
    }
    case '.markdown':
    case '.md': {
      return read_md_file({ file })
    }
    case '.json': {
      return [read_json_file({ file })]
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

export function write_file(args: {
  file: string
  sheets: SheetData<CellValue>[]
}) {
  let file = args.file
  let ext = extname(file)
  switch (ext) {
    case '.xlsx': {
      write_xlsx_file({ file, sheets: args.sheets })
      break
    }
    case '.csv':
    case '.tsv': {
      let separator = ext === '.tsv' ? '\t' : ','
      let files = infer_dest_files({ file, sheets: args.sheets })
      for (let i = 0; i < files.length; i++) {
        write_csv_file({
          file: files[i],
          rows: args.sheets[i].rows,
          separator,
        })
      }
      break
    }
    case '.md': {
      let files = infer_dest_files({ file, sheets: args.sheets })
      for (let i = 0; i < files.length; i++) {
        write_md_file({ file: files[i], rows: args.sheets[i].rows })
      }
      break
    }
    case '.json': {
      let files = infer_dest_files({ file, sheets: args.sheets })
      for (let i = 0; i < files.length; i++) {
        write_json_file({ file: files[i], rows: args.sheets[i].rows })
      }
      break
    }
    default: {
      throw new Error(`Unsupported file extension: ${ext}`)
    }
  }
}

function infer_dest_files(args: {
  file: string
  sheets: SheetData<CellValue>[]
}): string[] {
  let { file, sheets } = args
  if (sheets.length === 1) {
    return [file]
  }
  let name = infer_sheet_name(file)
  let ext = extname(file)
  return sheets.map(
    sheet => `${sanitize_name(name)}-${sanitize_name(sheet.name)}${ext}`,
  )
}

function sanitize_name(name: string): string {
  return name
    .replaceAll('/', '_')
    .replaceAll(':', '_')
    .replaceAll('\\', '_')
    .trim()
}

export function convert_file(args: { source_file: string; dest_file: string }) {
  let sheets = read_file({ file: args.source_file })
  write_file({ file: args.dest_file, sheets })
}

export type SheetInfo = {
  name: string
  range?: string
}

export function read_xlsx_file(args: {
  file: string
  sheets?: string[] | SheetInfo[]
}): SheetData<CellValue>[] {
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
    return { name: info.name, rows }
  })
}

export function read_csv_file(args: {
  file: string
  /** ',' | '\t' */
  separator?: string
}): SheetData<string> {
  let file = args.file
  let text = readFileSync(file, 'utf-8')

  let separator = args.separator
  if (!separator && file.endsWith('.tsv')) {
    separator = '\t'
  }
  if (!separator) {
    separator = infer_csv_separator(text)
  }

  let rows = from_csv(text, separator)
  let name = infer_sheet_name(file)
  return { name, rows }
}

export function infer_csv_separator(text: string) {
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

  return tab_count > comma_count ? '\t' : ','
}

export function read_md_file(args: { file: string }): SheetData<string>[] {
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
        cells.push(buffer.trim())
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
  let text = to_csv(
    args.rows.map(row => row.map(value_to_string)),
    separator || ',',
  )
  writeFileSync(file, text, 'utf-8')
}

function value_to_string(value: CellValue) {
  if (value instanceof Date) {
    return value.toISOString()
  }
  if (value === null) {
    return ''
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
  let rows = args.rows.map(cols => cols.map(value_to_string))

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
      let value = value_to_string(row[c])
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

export function read_json_file(args: {
  file: string
  format?: 'array' | 'object'
}): SheetData<CellValue> {
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
      return { name, rows: json }
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
          return read_json_file({ file, format: 'array' })
        }
        if (is_object(item)) {
          return read_json_file({ file, format: 'object' })
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
  format?: 'array' | 'object'
}) {
  let { file, rows } = args
  let values = args.format === 'array' ? rows : rows_to_objects(rows)
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
