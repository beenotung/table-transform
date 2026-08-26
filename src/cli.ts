import { readFileSync } from 'fs'
import { join } from 'path'
import {
  is_gsheets_ref,
  list_gsheets_names,
  list_sheet_names,
  parse_gsheets_ref,
  read_file,
  read_gsheets_file,
  resolve_sa_cred_file,
  write_file,
  write_gsheets_file,
} from './core'

let formats = ['md', 'markdown', 'csv', 'tsv', 'txt', 'xlsx', 'json']

type ShowName = 'auto' | 'always' | 'never'

function parse_column_names(args: {
  fields: string[]
  value: string | undefined
  label: string
}): void {
  let { fields, value, label } = args
  value = (value || '').trim()
  if (!value) {
    throw new Error(`Missing ${label}`)
  }
  for (let name of value.split(',')) {
    name = name.trim()
    if (name.length > 0) {
      fields.push(name)
    }
  }
}

function get_args() {
  let args = process.argv.slice(2)
  let input = ''
  let output = ''
  let format = ''
  let trim_string = true
  let trim_rows = true
  let trim_cols = true
  let separator = ''
  let input_separator = ''
  let output_separator = ''
  let json_format: 'object' | 'array' = 'object'
  let show_name: ShowName = 'auto'
  let list = false
  let sa_cred_file = ''
  let fields: string[] = []
  let exclude_fields: string[] = []
  let rest: string[] = []
  for (let i = 0; i < args.length; i++) {
    let arg = args[i]
    switch (arg) {
      case '-l':
      case '--list': {
        list = true
        break
      }
      case '-i':
      case '--input': {
        i++
        input = args[i]
        if (!input) {
          throw new Error('Missing input file')
        }
        break
      }
      case '-o':
      case '--output': {
        i++
        output = args[i]
        if (!output) {
          throw new Error('Missing output file')
        }
        break
      }
      case '--sa-cred-file': {
        i++
        sa_cred_file = args[i]
        if (!sa_cred_file) {
          throw new Error('Missing service account credential file')
        }
        break
      }
      case '--no-trim-string': {
        trim_string = false
        break
      }
      case '--no-trim-rows': {
        trim_rows = false
        break
      }
      case '--no-trim-cols': {
        trim_cols = false
        break
      }
      case '-s':
      case '--separator': {
        i++
        separator = args[i]
        if (!separator) {
          throw new Error('Missing separator')
        }
        break
      }
      case '--input-separator': {
        i++
        input_separator = args[i]
        if (!input_separator) {
          throw new Error('Missing input separator')
        }
        break
      }
      case '--output-separator': {
        i++
        output_separator = args[i]
        if (!output_separator) {
          throw new Error('Missing output separator')
        }
        break
      }
      case '--array': {
        json_format = 'array'
        break
      }
      case '--object': {
        json_format = 'object'
        break
      }
      case '-f':
      case '--format': {
        i++
        format = args[i]
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
      case '-c':
      case '--column':
      case '--columns':
      case '--field':
      case '--fields': {
        i++
        parse_column_names({
          fields,
          value: args[i],
          label: 'column names to include',
        })
        break
      }
      case '-x':
      case '--exclude':
      case '--exclude-column':
      case '--exclude-columns':
      case '--exclude-field':
      case '--exclude-fields': {
        i++
        parse_column_names({
          fields: exclude_fields,
          value: args[i],
          label: 'column names to exclude',
        })
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
      case 'version':
      case '-v':
      case '--version': {
        show_version()
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
  if (list) {
    if (output) {
      throw new Error('--list cannot be used with an output file')
    }
    if (format) {
      throw new Error('--list cannot be used with --format')
    }
    require_sa_cred_file({ files: [input], sa_cred_file })
    return {
      input,
      sa_cred_file,
      list: true as const,
    }
  }
  if (!output) {
    output = '/dev/stdout'
  }
  require_sa_cred_file({ files: [input, output], sa_cred_file })
  if (format && output !== '/dev/stdout') {
    throw new Error('format specified but output is not stdout')
  }
  if (!format && output === '/dev/stdout') {
    format = 'txt'
  }
  if (fields.length && exclude_fields.length) {
    throw new Error('Cannot use both column include and exclude options')
  }
  input_separator ||= separator
  output_separator ||= separator
  return {
    input,
    output,
    trim_string,
    trim_rows,
    trim_cols,
    input_separator,
    output_separator,
    json_format,
    format,
    show_name,
    fields,
    exclude_fields,
    sa_cred_file,
    list: false as const,
  }
}

/** fail before any network call when a gsheets ref has no credentials */
function require_sa_cred_file(args: {
  files: string[]
  sa_cred_file: string
}): void {
  // parsed first, so a malformed reference reports itself instead of the
  // credentials being blamed for it
  let has_gsheets = false
  for (let file of args.files) {
    if (parse_gsheets_ref(file)) {
      has_gsheets = true
    }
  }
  if (has_gsheets) {
    resolve_sa_cred_file({ sa_cred_file: args.sa_cred_file })
  }
}

function show_version() {
  let file = join(__dirname, '..', 'package.json')
  let text = readFileSync(file, 'utf-8')
  let pkg = JSON.parse(text)
  console.log(pkg.version)
}

function show_help() {
  console.log(
    `
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
  -s, --separator <char>     Example: '|' (default auto detect ',' or '\\t')
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
  - txt (ASCII box format from SQLite \`.mode table\`)
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

Examples:

  # list sheet/table names
  table-transform --list source.xlsx
  table-transform multi-table.md -l

  # convert file format
  table-transform source.xlsx export.csv
  export-table --output export.json source.md
  convert-table export.json --input source.md
  table-transform --no-trim-string source.xlsx export.json --array

  # output to console
  extract-table source.xlsx /dev/stdout --format csv
  extract-table source.xlsx --format csv
  extract-table source.xlsx - csv
  extract-table source.xlsx csv
  extract-table json source.xlsx
  extract-table source.xlsx

  # pick selected columns (instead of all columns)
  table-transform --column "Name, Tel" source.csv export.csv
  table-transform source.md -c "Name,Tel" export.csv
  table-transform source.csv -c Name -c Tel export.csv

  # exclude selected columns (include all other columns)
  table-transform --exclude "Email,CV" source.csv export.csv
  table-transform source.csv -x "Email, CV" export.csv

  # google sheets
  table-transform --list gsheets:1AbC --sa-cred-file service-account.json
  table-transform gsheets:1AbC/Roster export.csv --sa-cred-file service-account.json
  table-transform gsheets:1AbC#0 export.json --sa-cred-file service-account.json
  table-transform source.csv gsheets:1AbC/Roster --sa-cred-file service-account.json
`.trimStart(),
  )
}

async function main() {
  let args
  try {
    args = get_args()
  } catch (error) {
    console.error(String(error))
    process.exit(1)
  }
  if (args.list) {
    let names = is_gsheets_ref(args.input)
      ? await list_gsheets_names({
          file: args.input,
          sa_cred_file: args.sa_cred_file,
        })
      : list_sheet_names(args.input)
    for (let name of names) {
      console.log(name)
    }
    return
  }
  let output =
    args.output === '/dev/stdout' ? '/dev/stdout.' + args.format : args.output
  let read_options = {
    file: args.input,
    separator: args.input_separator,
    trim_string: args.trim_string,
    trim_rows: args.trim_rows,
    trim_cols: args.trim_cols,
    fields: args.fields,
    exclude_fields: args.exclude_fields,
  }
  let sheets = is_gsheets_ref(args.input)
    ? await read_gsheets_file({
        ...read_options,
        sa_cred_file: args.sa_cred_file,
      })
    : read_file(read_options)
  if (is_gsheets_ref(output)) {
    await write_gsheets_file({
      file: output,
      sheets,
      sa_cred_file: args.sa_cred_file,
    })
  } else {
    write_file({
      file: output,
      sheets,
      show_name: args.show_name,
      separator: args.output_separator,
      json_format: args.json_format,
    })
  }
}
main().catch(error => {
  console.error(String(error))
  process.exit(1)
})
