import { createSign } from 'crypto'
import { readFileSync, writeFileSync } from 'fs'

export type ServiceAccountCredentials = {
  client_email: string
  private_key: string
  token_uri?: string
}

export function extract_google_sheet_id(url_or_id: string): string {
  let match = url_or_id.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)
  return match ? match[1] : url_or_id.trim()
}

export function is_google_sheet_url(input: string): boolean {
  return /^https?:\/\/docs\.google\.com\/spreadsheets\/d\/[a-zA-Z0-9-_]+/.test(
    input.trim(),
  )
}

function base64url(input: string | Buffer): string {
  return Buffer.from(input)
    .toString('base64')
    .replaceAll('+', '-')
    .replaceAll('/', '_')
    .replace(/=+$/, '')
}

async function get_service_account_token(
  credentials: ServiceAccountCredentials,
): Promise<string> {
  let token_uri = credentials.token_uri || 'https://oauth2.googleapis.com/token'
  let now = Math.floor(Date.now() / 1000)

  let header = base64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }))
  let payload = base64url(
    JSON.stringify({
      iss: credentials.client_email,
      scope: 'https://www.googleapis.com/auth/drive.readonly',
      aud: token_uri,
      iat: now,
      exp: now + 3600,
    }),
  )

  let unsigned = `${header}.${payload}`
  let signer = createSign('RSA-SHA256')
  signer.update(unsigned)
  let signature = base64url(signer.sign(credentials.private_key))
  let jwt = `${unsigned}.${signature}`

  let res = await fetch(token_uri, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt,
    }),
  })

  if (!res.ok) {
    let detail = await res.text()
    throw new Error(
      `Failed to authenticate with Google Service Account (${res.status}): ${detail}`,
    )
  }

  let data = (await res.json()) as { access_token?: string }
  if (!data.access_token) {
    throw new Error('Missing access_token in Google OAuth response')
  }
  return data.access_token
}

function load_credentials(
  service_account?: string | ServiceAccountCredentials,
): ServiceAccountCredentials | undefined {
  let source = service_account || process.env.GOOGLE_APPLICATION_CREDENTIALS
  if (!source) {
    return undefined
  }
  if (typeof source === 'string') {
    let text = readFileSync(source, 'utf-8')
    return JSON.parse(text) as ServiceAccountCredentials
  }
  return source
}

export async function download_google_sheet(args: {
  /** Google Sheets URL or spreadsheet ID */
  url: string
  /** Destination local .xlsx file path */
  output_file: string
  /** Path to service account JSON file, or parsed credentials object (falls back to GOOGLE_APPLICATION_CREDENTIALS) */
  service_account?: string | ServiceAccountCredentials
}): Promise<string> {
  let spreadsheet_id = extract_google_sheet_id(args.url)
  if (!spreadsheet_id) {
    throw new Error(`Invalid Google Sheet URL or ID: ${args.url}`)
  }

  let credentials = load_credentials(args.service_account)

  let res: Response
  if (credentials) {
    let token = await get_service_account_token(credentials)
    let mime_type = encodeURIComponent(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    let export_url = `https://www.googleapis.com/drive/v3/files/${spreadsheet_id}/export?mimeType=${mime_type}`
    res = await fetch(export_url, {
      headers: { Authorization: `Bearer ${token}` },
      redirect: 'follow',
    })
  } else {
    let export_url = `https://docs.google.com/spreadsheets/d/${spreadsheet_id}/export?format=xlsx`
    res = await fetch(export_url, { redirect: 'follow' })
  }

  let content_type = res.headers.get('content-type') || ''
  if (!res.ok || content_type.includes('text/html')) {
    if (!credentials) {
      throw new Error(
        `Failed to download Google Sheet (HTTP ${res.status}). If the sheet is private, provide a service account JSON file via --service-account or GOOGLE_APPLICATION_CREDENTIALS, and share the sheet with its client_email.`,
      )
    }
    let detail = await res.text()
    throw new Error(
      `Failed to download Google Sheet with service account (HTTP ${res.status}): ${detail}`,
    )
  }

  let array_buffer = await res.arrayBuffer()
  writeFileSync(args.output_file, Buffer.from(array_buffer))
  return args.output_file
}
