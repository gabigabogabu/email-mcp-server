# Email MCP Server

A Model Context Protocol (MCP) server for email integration using IMAP and SMTP protocols. This server allows model assistants to access and interact with email accounts.

## Features

- Read emails from IMAP mailboxes
- Send emails via SMTP
- Search emails
- List email folders

## Prerequisites

- Node.js (v16 or higher)
- npm or yarn
- Email account with IMAP and SMTP access

## Setup

1. Clone the repository:
   ```
   git clone <repository-url>
   cd imap-mcp
   ```

2. Install dependencies:
   ```
   npm install
   ```

3. Create a `.env` file in the root directory with the following variables:
   ```
   EMAIL_USER=your-email@example.com
   EMAIL_PASSWORD=your-password
   IMAP_HOST=imap.example.com
   IMAP_PORT=993
   SMTP_HOST=smtp.example.com
   SMTP_PORT=465
   ```

## Usage

### Development

To run the server in development mode:

```
npm run dev
```

This will watch for changes, recompile TypeScript, and restart the server automatically.

### Production

To build and run the server in production mode:

```
npm run build
npm start
```

## API

The server exposes the following MCP resources:

- `mailto:<email-address>/inbox` - List of 10 most recent emails in the inbox
- `mailto:<email-address>/folders` - List of email folders/mailboxes

And the following MCP tools:

- `send_email` - Send an email message
- `search_emails` - Search for emails with advanced query options
- `list_folders` - List all available email folders/mailboxes

## License

ISC 