import { csv_to_json, from_csv } from '@beenotung/tslib'
import {
  convert_file,
  list_sheet_names,
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
import { readFileSync } from 'fs'

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

function test_newline() {
  convert_file({
    source_file: 'res/multi-line.csv',
    dest_file: 'res/multi-line.md',
  })
  let text = readFileSync('res/multi-line.md', 'utf-8')
  assert(
    text.includes('this is good<br>Line 2'),
    'failed to escape newline in md',
  )

  convert_file({
    source_file: 'res/multi-line.md',
    dest_file: 'res/multi-line.json',
  })
  text = readFileSync('res/multi-line.json', 'utf-8')
  let json = JSON.parse(text)
  assert(
    json[0].desc === 'this is good\nLine 2',
    'failed to unescape newline in md',
  )
}

function test_txt_file() {
  convert_file({
    source_file: 'res/table.md',
    dest_file: 'res/table-2.txt',
  })

  let expected = readFileSync('res/table.txt', 'utf-8').trim()
  let index = expected.indexOf('+-')
  expected = expected.slice(index)

  let actual = readFileSync('res/table-2.txt', 'utf-8').trim()
  assert(actual === expected, 'failed to output txt format')
}

function test_list_sheet_names() {
  assert.deepStrictEqual(list_sheet_names('res/multi-table.xlsx'), [
    'Users',
    'Posts',
  ])
  assert.deepStrictEqual(list_sheet_names('res/multi-table.md'), [
    'multi-table-1',
    'multi-table-2',
  ])
  assert.deepStrictEqual(list_sheet_names('res/roster.csv'), ['roster'])
}

// let data = test_xlsx_file()
// let data = test_md_file()
// test_convert()
test_convert_2()
// let data = test_multi_table()
// test_json_file()
// test_xlsx_range()
test_newline()
test_txt_file()
test_list_sheet_names()
// console.log(data)
// debugger

// let data = read_md_file({ file: 'res/roster.md' })

console.log('done.')
