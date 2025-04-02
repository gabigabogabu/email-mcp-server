#!/usr/bin/env node

import dotenv from 'dotenv';
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { ImapFlow, MessageAddressObject, FetchQueryObject } from 'imapflow';
import nodemailer from 'nodemailer';
import { z } from 'zod';

class InvalidEnvError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'InvalidEnvError';
  }
}

const parseEnv = () => {
  dotenv.config();
  const envParseResult = z.object({
    EMAIL_USER: z.string().email(),
    EMAIL_PASSWORD: z.string().min(1),
    IMAP_HOST: z.string().min(1),
    IMAP_PORT: z.string().regex(/^\d+$/).transform(Number),
    SMTP_HOST: z.string().min(1),
    SMTP_PORT: z.string().regex(/^\d+$/).transform(Number),
  }).safeParse(process.env);
  if (!envParseResult.success) throw new InvalidEnvError('Environment validation failed:');
  return envParseResult.data;
}
const env = parseEnv();

type EmailAddress = {
  name?: string;
  address: string;
}

type EmailMessage = {
  id: number;
  subject: string;
  from: EmailAddress[];
  to: EmailAddress[];
  date: Date;
  text?: string;
  html?: string;
}

type EmailFolder = {
  name: string;
  path: string;
  specialUse?: string | null;
  flags: string[];
}

// Initialize MCP server
const server = new McpServer({
  name: "imap-mcp",
  version: "0.1.0",
  capabilities: {
    resources: {},
    tools: {},
  },
});

// Set up email configuration
const emailUser = env.EMAIL_USER;
const emailConfig = {
  imap: {
    host: env.IMAP_HOST,
    port: env.IMAP_PORT,
    secure: true,
    auth: {
      user: emailUser,
      pass: env.EMAIL_PASSWORD
    }
  },
  smtp: {
    host: env.SMTP_HOST,
    port: env.SMTP_PORT,
    secure: env.SMTP_PORT === 465,
    auth: {
      user: emailUser,
      pass: env.EMAIL_PASSWORD
    }
  }
};

// Create IMAP client and SMTP transporter functions (to create fresh connections when needed)
const createImapClient = (): ImapFlow => new ImapFlow(emailConfig.imap);
const smtpTransporter = nodemailer.createTransport(emailConfig.smtp);

const mapAddress = (addr: MessageAddressObject): EmailAddress => ({
  name: addr.name,
  address: addr.address || ''
})

// Tool definitions
server.tool(
  "send_email",
  "Send an email message",
  {
    to: z.string().email(),
    subject: z.string().min(1),
    text: z.string().min(1),
    html: z.string().optional(),
  },
  async ({ to, subject, text, html }) => {
    try {
      const mailOptions = {
        from: emailUser,
        to,
        subject,
        text,
        html
      };

      const info = await smtpTransporter.sendMail(mailOptions);

      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            messageId: info.messageId,
            status: 'Email sent successfully'
          }, null, 2)
        }]
      };
    } catch (error) {
      console.error("Error sending email:", error);
      const err = error as Error;
      return {
        content: [{ type: "text", text: `Error sending email: ${err.message}` }],
        isError: true
      };
    }
  }
);

let searchQuerySchema = z.object({
  seq: z.string().optional(),
  answered: z.boolean().optional(),
  deleted: z.boolean().optional(),
  draft: z.boolean().optional(),
  flagged: z.boolean().optional(),
  seen: z.boolean().optional(),
  all: z.boolean().optional(),
  new: z.boolean().optional(),
  old: z.boolean().optional(),
  recent: z.boolean().optional(),
  from: z.string().optional(),
  to: z.string().optional(),
  cc: z.string().optional(),
  bcc: z.string().optional(),
  body: z.string().optional(),
  subject: z.string().min(1),
  larger: z.number().int().positive().optional(),
  smaller: z.number().int().positive().optional(),
  uid: z.string().optional(),
  modseq: z.bigint().optional(),
  emailId: z.string().optional(),
  threadId: z.string().optional(),
  before: z.date().optional(),
  on: z.date().optional(),
  since: z.date().optional(),
  sentBefore: z.date().optional(),
  sentOn: z.date().optional(),
  sentSince: z.date().optional(),
  keyword: z.string().optional(),
  unKeyword: z.string().optional(),
  header: z.record(z.string(), z.union([z.boolean(), z.string()])).optional(),
})
searchQuerySchema = searchQuerySchema.extend({
  // recursive OR queries
  or: z.array(searchQuerySchema).optional(),
})

const fetchOptionsSchema = z.object({
  uid: z.boolean().optional(),
  flags: z.boolean().optional(),
  bodyStructure: z.boolean().optional(),
  envelope: z.boolean().optional(),
  internalDate: z.boolean().optional(),
  size: z.boolean().optional(),
  source: z.union([z.boolean(), z.object({})]).optional(),
  threadId: z.boolean().optional(),
  labels: z.boolean().optional(),
  headers: z.union([z.boolean(), z.array(z.string())]).optional(),
  bodyParts: z.array(z.string()).optional(),
})

server.tool(
  "search_emails",
  "Search for emails in the inbox",
  {
    query: searchQuerySchema,
    fetchOptions: fetchOptionsSchema,
    folder: z.string().default("INBOX"),
    limit: z.number().int().positive().default(10),
  },
  async ({ query, folder, limit, fetchOptions }) => {
    const client = createImapClient();

    try {
      await client.connect();
      await client.mailboxOpen(folder);

      // Search for messages matching the query
      const uids = await client.search(query, { uid: true });

      const messages: EmailMessage[] = [];

      // Limit to the specified number of messages
      const limitedUids = uids.slice(0, limit);

      if (limitedUids.length > 0) {
        for await (const message of client.fetch(limitedUids, fetchOptions)) {
          messages.push({
            id: message.uid,
            subject: message.envelope.subject,
            from: message.envelope.from.map(mapAddress),
            to: message.envelope.to.map(mapAddress),
            date: message.envelope.date
          });
        }
      }

      return {
        content: [{ type: "text", text: JSON.stringify(messages, null, 2) }]
      };
    } catch (error) {
      console.error("Error searching emails:", error);
      const err = error as Error;
      return {
        content: [{ type: "text", text: `Error searching emails: ${err.message}` }],
        isError: true
      };
    } finally {
      await client.logout();
    }
  }
);

server.tool(
  "list_folders",
  "List all available email folders/mailboxes",
  {},
  async () => {
    const client = createImapClient();

    try {
      await client.connect();

      const mailboxes: EmailFolder[] = [];
      const listResponse = await client.list();

      for (const mailbox of listResponse) {
        mailboxes.push({
          name: mailbox.name,
          path: mailbox.path,
          specialUse: mailbox.specialUse,
          flags: Array.from(mailbox.flags)
        });
      }

      return {
        content: [{ type: "text", text: JSON.stringify(mailboxes, null, 2) }]
      };
    } catch (error) {
      console.error("Error listing folders:", error);
      const err = error as Error;
      return {
        content: [{ type: "text", text: `Error listing folders: ${err.message}` }],
        isError: true
      };
    } finally {
      await client.logout();
    }
  }
);

// Resource handlers
server.resource(
  "inbox",
  "mailto:" + emailUser + "/inbox",
  async () => {
    const client = createImapClient();

    try {
      await client.connect();
      await client.mailboxOpen('INBOX');

      // Get the 10 most recent messages
      const messages: EmailMessage[] = [];
      const fetchOptions: FetchQueryObject = { envelope: true };
      
      for await (const message of client.fetch('1:*', fetchOptions)) {
        messages.push({
          id: message.uid,
          subject: message.envelope.subject,
          from: message.envelope.from.map(mapAddress),
          to: message.envelope.to.map(mapAddress),
          date: message.envelope.date
        });
      }

      // Sort by date, newest first, and limit to 10
      messages.sort((a, b) => b.date.getTime() - a.date.getTime());
      const recentMessages = messages.slice(0, 10);

      return {
        contents: [{
          uri: "mailto:" + emailUser + "/inbox",
          mimeType: "application/json",
          text: JSON.stringify(recentMessages, null, 2)
        }]
      };
    } catch (error) {
      console.error("Error reading inbox:", error);
      throw error;
    } finally {
      await client.logout();
    }
  }
);

server.resource(
  "folders",
  "mailto:" + emailUser + "/folders",
  async () => {
    const client = createImapClient();

    try {
      await client.connect();
      const mailboxes: EmailFolder[] = [];
      const listResponse = await client.list();

      for (const mailbox of listResponse) {
        mailboxes.push({
          name: mailbox.name,
          path: mailbox.path,
          specialUse: mailbox.specialUse,
          flags: Array.from(mailbox.flags)
        });
      }

      return {
        contents: [{
          uri: "mailto:" + emailUser + "/folders",
          mimeType: "application/json",
          text: JSON.stringify(mailboxes, null, 2)
        }]
      };
    } catch (error) {
      console.error("Error listing folders:", error);
      throw error;
    } finally {
      await client.logout();
    }
  }
);

// Start the MCP server
async function runServer(): Promise<void> {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

runServer().catch(console.error); 