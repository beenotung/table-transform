import {
  CellValue,
  read_file,
  write_csv_file,
  write_file,
  write_json_file,
  write_md_file,
  write_xlsx_file,
} from './core'

let formats = ['md', 'markdown', 'csv', 'tsv', 'xlsx', 'json']

function get_args() {
  let args = process.argv.slice(2)
  let input = ''
  let output = ''
  let format = ''
  let rest: string[] = []
  for (let i = 0; i < args.length; i++) {
    let arg = args[i]
    switch (arg) {
      case '-i':
      case '--input': {
        i++
        input = arg
        if (!input) {
          throw new Error('Missing input file')
        }
        break
      }
      case '-o':
      case '--output': {
        i++
        output = arg
        if (!output) {
          throw new Error('Missing output file')
        }
        break
      }
      case '-f':
      case '--format': {
        i++
        format = arg
        if (!format) {
          throw new Error('Missing format')
        }
        if (!formats.includes(format)) {
          throw new Error(`Invalid format: ${format}`)
        }
        break
      }
      case '-h':
      case '--help': {
        show_help()
        process.exit(0)
      }
      case '-': {
        output = '/dev/stdout'
        break
      }
      default: {
        rest.push(arg)
      }
    }
  }
  for (let arg of rest) {
    if (formats.includes(arg)) {
      if (format) {
        throw new Error(`Multiple formats specified: ${format} and ${arg}`)
      }
      format = arg
      continue
    }
    if (!input) {
      input = arg
      continue
    }
    if (!output) {
      output = arg
      continue
    }
    throw new Error(`Unknown argument: ${arg}`)
  }
  if (!input) {
    throw new Error('Missing input file')
  }
  if (!output) {
    output = '/dev/stdout'
  }
  if (format && output !== '/dev/stdout') {
    throw new Error('format specified but output is not stdout')
  }
  if (!format && output === '/dev/stdout') {
    format = 'markdown'
  }
  return { input, output, format }
}

function show_help() {
  console.log(
    `
Usage:
  table-transform [options]

Alias: table-cli, table-convert

Options:
  -i, --input <file>     Input file (path)
  -o, --output <file>    Output file (path or /dev/stdout)
  -f, --format <format>  Output format (only for console output, default: markdown)
  -h, --help             Show help

Supported formats:
  - md, markdown
  - csv, tsv
  - xlsx
  - json

Examples:

  # convert file format
  table-transform source.xlsx export.csv
  table-convert --output export.json source.md
  table-convert export.json --input source.md

  # output to console
  table-cli source.xlsx /dev/stdout --format csv
  table-cli source.xlsx --format csv
  table-cli source.xlsx - csv
  table-cli source.xlsx csv
  table-cli json source.xlsx
  table-cli source.xlsx
`.trimStart(),
  )
}

function main() {
  let args
  try {
    args = get_args()
  } catch (error) {
    console.error(String(error))
    process.exit(1)
  }
  let sheets = read_file({ file: args.input })
  if (args.output !== '/dev/stdout') {
    write_file({ file: args.output, sheets })
    return
  }
  let file = '/dev/stdout'
  let write: (rows: CellValue[][]) => void
  switch (args.format) {
    case 'xlsx': {
      write_xlsx_file({ file, sheets })
      return
    }
    case 'csv': {
      write = rows => write_csv_file({ file, rows })
      break
    }
    case 'tsv': {
      write = rows => write_csv_file({ file, rows, separator: '\t' })
      break
    }
    case 'markdown':
    case 'md': {
      write = rows => write_md_file({ file, rows })
      break
    }
    case 'json': {
      write = rows => write_json_file({ file, rows })
      break
    }
    default: {
      throw new Error(`Unsupported format: ${args.format}`)
    }
  }
  let show_name = sheets.length > 1
  for (let sheet of sheets) {
    if (show_name) {
      console.log(`Sheet: ${sheet.name}`)
    }
    write(sheet.rows)
  }
}
main()
