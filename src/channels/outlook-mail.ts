/**
 * Outlook Mail channel — Microsoft Graph API (MSAL Device Code flow)
 *
 * Required .env variables:
 *   MICROSOFT_CLIENT_ID   — Azure App Registration client ID
 *   MICROSOFT_TENANT_ID   — Azure tenant ID (or "consumers" for personal accounts)
 *
 * On first run a device code URL + code will be printed. Open the URL in a browser,
 * enter the code and sign in. The refresh token is cached in ./store/ms-token.json.
 *
 * Microsoft Graph scopes needed:
 *   Mail.Read, Mail.ReadWrite, Mail.Send
 */

import {
  PublicClientApplication,
  type DeviceCodeRequest,
} from '@azure/msal-node';
import fs from 'fs';
import path from 'path';

import { logger } from '../logger.js';
import { readEnvFile } from '../env.js';
import { STORE_DIR, POLL_INTERVAL } from '../config.js';
import {
  Channel,
  NewMessage,
  OnInboundMessage,
  OnChatMetadata,
} from '../types.js';
import { registerChannel } from './registry.js';

const TOKEN_CACHE_PATH = path.join(STORE_DIR, 'ms-token.json');
const MAIL_POLL_MS = 60_000; // check for new mail every minute
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const SCOPES = [
  'Mail.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.Read',
  'Calendars.ReadWrite',
];

// ── Auth helpers ───────────────────────────────────────────────────────────

function buildMsalApp(
  clientId: string,
  tenantId: string,
): PublicClientApplication {
  const cachePlugin = {
    beforeCacheAccess: async (ctx: any) => {
      if (fs.existsSync(TOKEN_CACHE_PATH)) {
        ctx.tokenCache.deserialize(fs.readFileSync(TOKEN_CACHE_PATH, 'utf-8'));
      }
    },
    afterCacheAccess: async (ctx: any) => {
      if (ctx.cacheHasChanged) {
        fs.mkdirSync(path.dirname(TOKEN_CACHE_PATH), { recursive: true });
        fs.writeFileSync(TOKEN_CACHE_PATH, ctx.tokenCache.serialize());
      }
    },
  };

  return new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: { cachePlugin },
  });
}

async function getAccessToken(app: PublicClientApplication): Promise<string> {
  // Try silent first (uses cached refresh token)
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await app.acquireTokenSilent({
        account: accounts[0],
        scopes: SCOPES,
      });
      if (result?.accessToken) return result.accessToken;
    } catch {
      // Fall through to device code
    }
  }

  // Device code flow
  const request: DeviceCodeRequest = {
    scopes: SCOPES,
    deviceCodeCallback: (response: Record<string, unknown>) => {
      const msg = (response.message ??
        response.userCode ??
        JSON.stringify(response)) as string;
      process.stdout.write('\n=== Microsoft Authentication Required ===\n');
      process.stdout.write(msg + '\n');
      process.stdout.write('=========================================\n\n');
    },
  };
  const result = await app.acquireTokenByDeviceCode(request);
  if (!result?.accessToken) throw new Error('Microsoft auth failed');
  return result.accessToken;
}

// ── Graph API helpers ──────────────────────────────────────────────────────

async function graphGet(token: string, endpoint: string): Promise<any> {
  const res = await fetch(`${GRAPH_BASE}${endpoint}`, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
  });
  if (!res.ok)
    throw new Error(`Graph GET ${endpoint} → ${res.status} ${res.statusText}`);
  return res.json();
}

async function graphPost(
  token: string,
  endpoint: string,
  body: unknown,
): Promise<any> {
  const res = await fetch(`${GRAPH_BASE}${endpoint}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok)
    throw new Error(`Graph POST ${endpoint} → ${res.status} ${res.statusText}`);
  if (res.status === 202 || res.headers.get('content-length') === '0')
    return null;
  return res.json();
}

async function graphPatch(
  token: string,
  endpoint: string,
  body: unknown,
): Promise<void> {
  const res = await fetch(`${GRAPH_BASE}${endpoint}`, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok)
    throw new Error(
      `Graph PATCH ${endpoint} → ${res.status} ${res.statusText}`,
    );
}

/** Move a message to a named folder (looks up folder ID by display name). */
async function moveMessage(
  token: string,
  messageId: string,
  folderName: string,
): Promise<void> {
  const data = await graphGet(
    token,
    `/me/mailFolders?$filter=displayName eq '${encodeURIComponent(folderName)}'&$select=id,displayName`,
  );
  const folder = data.value?.[0];
  if (!folder) {
    throw new Error(`Outlook Mail: folder "${folderName}" not found`);
  }
  await graphPost(token, `/me/messages/${messageId}/move`, {
    destinationId: folder.id,
  });
  logger.info(
    { messageId, folderName },
    `Outlook Mail: moved message to "${folderName}"`,
  );
}

// ── Channel factory ────────────────────────────────────────────────────────

function createOutlookMailChannel(opts: {
  onMessage: OnInboundMessage;
  onChatMetadata: OnChatMetadata;
  registeredGroups: () => Record<string, import('../types.js').RegisteredGroup>;
}): Channel | null {
  const env = readEnvFile(['MICROSOFT_CLIENT_ID', 'MICROSOFT_TENANT_ID']);
  const clientId = process.env.MICROSOFT_CLIENT_ID || env.MICROSOFT_CLIENT_ID;
  const tenantId =
    process.env.MICROSOFT_TENANT_ID || env.MICROSOFT_TENANT_ID || 'consumers';

  if (!clientId) {
    logger.warn('Outlook Mail: MICROSOFT_CLIENT_ID not set — channel skipped');
    return null;
  }

  const msalApp = buildMsalApp(clientId, tenantId);
  let connected = false;
  let pollTimer: NodeJS.Timeout | null = null;
  let currentToken = '';

  async function refreshToken(): Promise<string> {
    currentToken = await getAccessToken(msalApp);
    return currentToken;
  }

  async function pollUnreadMails(): Promise<void> {
    try {
      const token = await refreshToken();
      const data = await graphGet(
        token,
        `/me/mailFolders/Inbox/messages?$filter=isRead eq false&$orderby=receivedDateTime desc&$top=20&$select=id,subject,bodyPreview,from,receivedDateTime,conversationId`,
      );

      for (const mail of data.value ?? []) {
        const jid = `mail:${mail.id}`;
        const senderAddress: string =
          mail.from?.emailAddress?.address ?? 'unknown';
        const senderName: string =
          mail.from?.emailAddress?.name ?? senderAddress;
        const subject: string = mail.subject ?? '(sans objet)';
        const preview: string = mail.bodyPreview ?? '';
        const received: string =
          mail.receivedDateTime ?? new Date().toISOString();

        const content = `📧 **${subject}**\nDe : ${senderName} <${senderAddress}>\n\n${preview}`;

        const newMsg: NewMessage = {
          id: mail.id,
          chat_jid: jid,
          sender: senderAddress,
          sender_name: senderName,
          content,
          timestamp: received,
          is_from_me: false,
        };

        opts.onChatMetadata(
          jid,
          received,
          `Mail: ${subject}`,
          'outlook-mail',
          false,
        );
        opts.onMessage(jid, newMsg);

        // Mark as read so it won't re-appear next poll
        await graphPatch(token, `/me/messages/${mail.id}`, { isRead: true });
      }
    } catch (err) {
      logger.error({ err }, 'Outlook Mail: poll failed');
    }
  }

  return {
    name: 'outlook-mail',

    async connect() {
      // Authenticate asynchronously so NanoClaw starts without blocking
      setImmediate(async () => {
        try {
          logger.info(
            'Outlook Mail: authenticating… (device code will appear in console)',
          );
          await refreshToken();
          connected = true;
          logger.info('Outlook Mail: connected ✓ — polling every 60s');
          await pollUnreadMails();
          pollTimer = setInterval(pollUnreadMails, MAIL_POLL_MS);
        } catch (err) {
          logger.error(
            { err },
            'Outlook Mail: authentication failed — restart after completing auth',
          );
        }
      });
    },

    async sendMessage(jid: string, text: string) {
      // jid format: "mail:<originalMessageId>"
      const token = await refreshToken();
      const originalId = jid.startsWith('mail:') ? jid.slice(5) : null;

      // Special command: "__MOVE__:<folderName>" — moves the message to a folder
      if (text.startsWith('__MOVE__:') && originalId) {
        const folderName = text.slice('__MOVE__:'.length).trim();
        await moveMessage(token, originalId, folderName);
        return;
      }

      if (originalId) {
        // Send as reply
        await graphPost(token, `/me/messages/${originalId}/reply`, {
          comment: text,
        });
      } else {
        // Standalone send — jid is an email address
        await graphPost(token, `/me/sendMail`, {
          message: {
            subject: 'Message de votre assistant',
            body: { contentType: 'Text', content: text },
            toRecipients: [{ emailAddress: { address: jid } }],
          },
        });
      }
    },

    isConnected() {
      return connected;
    },

    ownsJid(jid: string) {
      return (
        jid.startsWith('mail:') ||
        (jid.includes('@') &&
          !jid.includes('@s.whatsapp.net') &&
          !jid.endsWith('@g.us'))
      );
    },

    async disconnect() {
      if (pollTimer) clearInterval(pollTimer);
      connected = false;
    },
  };
}

registerChannel('outlook-mail', createOutlookMailChannel);
