import { csv_to_json, from_csv } from '@beenotung/tslib'
import {
  convert_file,
  read_csv_file,
  read_file,
  read_json_file,
  read_md_file,
  read_xlsx_file,
  rows_to_objects,
  write_csv_file,
  write_file,
  write_json_file,
  write_md_file,
  write_xlsx_file,
} from './core'
import assert from 'assert'

function test_xlsx_file() {
  let rows = read_xlsx_file({
    file: 'res/Attendance.xlsx',
    sheets: [
      { name: '點名紙', range: 'A14:E26' },
      { name: '聯絡表', range: 'A10:F22' },
      { name: '教師簽到表', range: 'A11:F19' },
    ],
  }).find(sheet => sheet.name === '點名紙')!.rows
  rows[0] = rows[0].map(label => {
    label = String(label).trim()
    if (label.includes('English')) return 'en'
    if (label.includes('Chi')) return 'zh'
    if (label.includes('Nickname')) return 'nickname'
    if (label.includes('Student No')) return 'id'
    if (label.includes('No')) return 'no'
    return label
  })
  return csv_to_json(rows as string[][]).filter(item => item.id)
}

function test_md_file() {
  let rows = read_md_file({ file: 'res/roster.md' })[0].rows
  // let rows = read_md_file({ file: 'res/escape.md' })
  rows[0] = rows[0].map(label => {
    if (label.includes('Eng')) return 'en'
    if (label.includes('Chi')) return 'zh'
    if (label.includes('SID')) return 'id'
    if (label.includes('No')) return 'no'
    if (label.includes('Status')) return 'status'
    return label
  })
  return csv_to_json(rows as string[][])
}

function test_convert() {
  write_csv_file({
    file: 'res/roster.csv',
    rows: read_md_file({ file: 'res/roster.md' })[0].rows,
  })
  write_csv_file({
    file: 'res/roster.tsv',
    rows: read_csv_file({ file: 'res/roster.csv' }).rows,
  })
  write_xlsx_file({
    file: 'res/roster.xlsx',
    sheets: [read_csv_file({ file: 'res/roster.tsv' })],
  })
  write_md_file({
    file: 'res/roster-2.md',
    rows: read_xlsx_file({ file: 'res/roster.xlsx' })[0].rows,
  })
}

function test_convert_2() {
  convert_file({ source_file: 'res/roster.md', dest_file: 'res/roster.csv' })
  convert_file({ source_file: 'res/roster.csv', dest_file: 'res/roster.tsv' })
  convert_file({ source_file: 'res/roster.tsv', dest_file: 'res/roster.xlsx' })
  convert_file({ source_file: 'res/roster.xlsx', dest_file: 'res/roster.json' })
  convert_file({
    source_file: 'res/roster.json',
    dest_file: 'res/roster.txt',
    separator: ' | ',
  })
  convert_file({
    source_file: 'res/roster.txt',
    separator: '|',
    dest_file: 'res/roster-2.md',
  })
}

function test_type_infer() {
  read_file({ file: 'res/data.csv' })[0].rows[0][0].toUpperCase()
  read_file({ file: 'res/data.tsv' })[0].rows[0][0].toUpperCase()
  read_file({ file: 'res/data.md' })[0].rows[0][0].toUpperCase()
  // @ts-expect-error
  read_file({ file: 'res/data.xlsx' })[0].rows[0][0].toUpperCase()
}

function test_multi_table() {
  let sheets = read_md_file({ file: 'res/multi-table.md' })
  sheets[0].name = 'Users'
  sheets[1].name = 'Posts'
  write_file({
    file: 'res/multi-table.xlsx',
    sheets: sheets,
  })
  let data = sheets.map(sheet => ({
    name: sheet.name,
    data: rows_to_objects(sheet.rows),
  }))
  return data
}

function test_json_file() {
  let rows = read_json_file({ file: 'res/roster.json' }).rows
  let text = JSON.stringify(rows)
  write_json_file({ file: 'res/roster-array.json', rows, format: 'array' })
  write_json_file({ file: 'res/roster-object.json', rows, format: 'object' })
  assert(
    text ===
      JSON.stringify(read_json_file({ file: 'res/roster-array.json' }).rows),
    'failed to read json array of arrays',
  )
  assert(
    text ===
      JSON.stringify(read_json_file({ file: 'res/roster-object.json' }).rows),
    'failed to read json array of objects',
  )
}

function test_xlsx_range() {
  let file = 'res/with-empty-cells.xlsx'
  let data = read_xlsx_file({ file })
  write_json_file({ file: 'test.json', rows: data[0].rows, format: 'array' })
}

// let data = test_xlsx_file()
// let data = test_md_file()
// test_convert()
test_convert_2()
// let data = test_multi_table()
// test_json_file()
// test_xlsx_range()
// console.log(data)
// debugger

// let data = read_md_file({ file: 'res/roster.md' })

console.log('done.')
