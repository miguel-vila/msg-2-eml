# msg-2-eml

A web-based tool to convert Microsoft Outlook MSG files to standard EML format.

## Features

- Drag-and-drop or click-to-upload interface
- Batch conversion of multiple MSG files
- Preserves email metadata (From, To, Cc, Bcc, Subject, Date)
- Supports plain text, HTML, and RTF email bodies
- Handles attachments including embedded/forwarded emails
- Calendar event support with iCalendar generation
- RFC-compliant output:
  - RFC 2047 encoding for non-ASCII subjects and display names
  - RFC 2231 encoding for non-ASCII attachment filenames
  - RFC 5322 header folding for long headers
- Preserves read receipt and delivery receipt request headers

## Requirements

- Node.js 18+

## Installation

```bash
npm install
```

## Usage

### Development

```bash
npm run dev
```

This builds the project and starts the server at http://localhost:3000.

### Production

```bash
npm run build
npm start
```

### Running Tests

```bash
npm test
```

## API

### `POST /api/convert`

Converts a single MSG file to EML format.

**Request:**
- Body: Raw MSG file binary data
- Content-Type: `application/octet-stream`

**Response:**
- Content-Type: `message/rfc822`
- Body: Converted EML file

### `GET /api/health`

Health check endpoint.

## Project Structure

```
msg-2-eml/
├── src/
│   ├── client/          # Frontend web app
│   │   ├── app.ts
│   │   └── index.html
│   └── server/          # Express backend
│       ├── index.ts
│       └── msg-to-eml.ts
├── dist/                # Compiled output
├── package.json
└── tsconfig.json
```

## How It Works

1. MSG files use Microsoft's OLE Compound Document format
2. The `msg-parser` library extracts email properties and attachments
3. Compressed RTF bodies are decompressed and de-encapsulated to extract HTML/text
4. The extracted data is assembled into a MIME-compliant EML structure
5. Attachments are base64 encoded; embedded messages are recursively converted

## License

MIT
