import { read_file, write_file } from './core'

let formats = ['md', 'markdown', 'csv', 'tsv', 'xlsx', 'json']

type ShowName = 'auto' | 'always' | 'never'

function get_args() {
  let args = process.argv.slice(2)
  let input = ''
  let output = ''
  let format = ''
  let show_name: ShowName = 'auto'
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
      case '-n':
      case '--name': {
        i++
        arg = args[i]
        switch (arg) {
          case 'auto':
          case 'always':
          case 'never': {
            show_name = arg
            break
          }
          default: {
            if (arg) {
              throw new Error(`Invalid table name option: ${arg}`)
            } else {
              throw new Error('Missing show name')
            }
          }
        }
        break
      }
      case '--show-name': {
        show_name = 'always'
        break
      }
      case '--hide-name': {
        show_name = 'never'
        break
      }
      case 'auto':
      case 'always':
      case 'never': {
        show_name = arg
        break
      }
      case 'help':
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
  return { input, output, format, show_name }
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
  -h, --help             Show help

Options only for console output:
  -f, --format <format>  Output format (default: markdown)
  -n, --name <mode>      Display table name or not (default: auto)
  --show-name                     Alias for "--name always"
  --hide-name                     Alias for "--name never"

Supported formats:
  - md, markdown
  - csv, tsv
  - xlsx
  - json

Display table name modes:
  auto   : display table name only if there are multiple tables
  always : always display table name
  never  : never display table name

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
  let output =
    args.output === '/dev/stdout' ? '/dev/stdout.' + args.format : args.output
  let sheets = read_file({ file: args.input })
  write_file({ file: output, sheets, show_name: args.show_name })
}
main()
