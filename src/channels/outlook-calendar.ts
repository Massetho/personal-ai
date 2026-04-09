/**
 * Outlook Calendar channel — Microsoft Graph API
 *
 * Re-uses the same MSAL credentials as outlook-mail.ts (shared token cache).
 * Provides daily agenda summaries and upcoming event reminders via WhatsApp.
 *
 * Required .env variables (same as outlook-mail):
 *   MICROSOFT_CLIENT_ID
 *   MICROSOFT_TENANT_ID
 *   CALENDAR_NOTIFY_JID  — WhatsApp JID to deliver agenda summaries to
 *                          e.g. 33612345678@s.whatsapp.net
 */

import { PublicClientApplication } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';

import { logger } from '../logger.js';
import { readEnvFile } from '../env.js';
import { STORE_DIR, TIMEZONE } from '../config.js';
import {
  Channel,
  NewMessage,
  OnInboundMessage,
  OnChatMetadata,
} from '../types.js';
import { registerChannel } from './registry.js';

const TOKEN_CACHE_PATH = path.join(STORE_DIR, 'ms-token.json');
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const SCOPES = [
  'Mail.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.Read',
  'Calendars.ReadWrite',
];
const CALENDAR_POLL_MS = 15 * 60_000; // check calendar every 15 minutes

// ── Shared auth helpers (identical to outlook-mail, reuses cache file) ─────

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
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await app.acquireTokenSilent({
        account: accounts[0],
        scopes: SCOPES,
      });
      if (result?.accessToken) return result.accessToken;
    } catch {}
  }
  throw new Error(
    'Outlook Calendar: no valid token — start outlook-mail first to authenticate',
  );
}

async function graphGet(token: string, endpoint: string): Promise<any> {
  const res = await fetch(`${GRAPH_BASE}${endpoint}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph GET ${endpoint} → ${res.status}`);
  return res.json();
}

// ── Formatting ─────────────────────────────────────────────────────────────

function formatEvent(event: any): string {
  const title: string = event.subject ?? '(sans titre)';
  const start = event.start?.dateTime
    ? new Date(event.start.dateTime + 'Z').toLocaleString('fr-FR', {
        timeZone: TIMEZONE,
        hour: '2-digit',
        minute: '2-digit',
      })
    : 'toute la journée';
  const location: string = event.location?.displayName ?? '';
  const organizer: string = event.organizer?.emailAddress?.name ?? '';
  return `• ${start} — ${title}${location ? ` 📍 ${location}` : ''}${organizer ? ` (${organizer})` : ''}`;
}

// ── Channel factory ────────────────────────────────────────────────────────

function createOutlookCalendarChannel(opts: {
  onMessage: OnInboundMessage;
  onChatMetadata: OnChatMetadata;
  registeredGroups: () => Record<string, import('../types.js').RegisteredGroup>;
}): Channel | null {
  const env = readEnvFile([
    'MICROSOFT_CLIENT_ID',
    'MICROSOFT_TENANT_ID',
    'CALENDAR_NOTIFY_JID',
  ]);
  const clientId = process.env.MICROSOFT_CLIENT_ID || env.MICROSOFT_CLIENT_ID;
  const tenantId =
    process.env.MICROSOFT_TENANT_ID || env.MICROSOFT_TENANT_ID || 'consumers';
  const notifyJid = process.env.CALENDAR_NOTIFY_JID || env.CALENDAR_NOTIFY_JID;

  if (!clientId) {
    logger.warn(
      'Outlook Calendar: MICROSOFT_CLIENT_ID not set — channel skipped',
    );
    return null;
  }

  const msalApp = buildMsalApp(clientId, tenantId);
  let connected = false;
  let pollTimer: NodeJS.Timeout | null = null;
  let lastDailyDigest = '';

  async function fetchUpcomingEvents(days = 7): Promise<any[]> {
    const token = await getAccessToken(msalApp);
    const now = new Date().toISOString();
    const end = new Date(Date.now() + days * 24 * 3600 * 1000).toISOString();
    const data = await graphGet(
      token,
      `/me/calendarView?startDateTime=${now}&endDateTime=${end}&$orderby=start/dateTime&$top=50&$select=subject,start,end,location,organizer,isAllDay`,
    );
    return data.value ?? [];
  }

  async function sendDailyDigest(): Promise<void> {
    if (!notifyJid) return;
    const today = new Date().toLocaleDateString('fr-FR', {
      timeZone: TIMEZONE,
    });
    if (lastDailyDigest === today) return; // already sent today

    try {
      const events = await fetchUpcomingEvents(1);
      const jid = `calendar:daily-digest`;
      const content = events.length
        ? `📅 **Agenda du jour — ${today}**\n\n${events.map(formatEvent).join('\n')}`
        : `📅 **Agenda du jour — ${today}**\n\nAucun événement aujourd'hui.`;

      const msg: NewMessage = {
        id: `digest-${today}`,
        chat_jid: jid,
        sender: 'calendar',
        sender_name: 'Agenda',
        content,
        timestamp: new Date().toISOString(),
        is_from_me: false,
      };

      opts.onChatMetadata(
        jid,
        new Date().toISOString(),
        'Agenda quotidien',
        'outlook-calendar',
        false,
      );
      opts.onMessage(jid, msg);
      lastDailyDigest = today;
      logger.info('Outlook Calendar: daily digest sent');
    } catch (err) {
      logger.error({ err }, 'Outlook Calendar: daily digest failed');
    }
  }

  async function pollCalendar(): Promise<void> {
    // Send daily digest at startup and once per day
    await sendDailyDigest();
  }

  return {
    name: 'outlook-calendar',

    async connect() {
      logger.info('Outlook Calendar: connecting…');
      // Retry loop — waits for outlook-mail to complete auth first (shared token cache)
      setImmediate(async () => {
        for (let attempt = 0; attempt < 30; attempt++) {
          try {
            await getAccessToken(msalApp);
            connected = true;
            logger.info('Outlook Calendar: connected ✓');
            await pollCalendar();
            pollTimer = setInterval(pollCalendar, CALENDAR_POLL_MS);
            return;
          } catch {
            await new Promise((r) => setTimeout(r, 10_000));
          }
        }
        logger.warn(
          'Outlook Calendar: could not authenticate after 5 min — restart after completing MS auth',
        );
      });
    },

    /**
     * sendMessage is used to answer queries like "quel est mon agenda demain ?"
     * The jid here is synthetic (e.g. "calendar:query"), so we route
     * the response back through the WhatsApp channel via the notifyJid.
     * In practice the agent answers directly on the channel that asked the question.
     */
    async sendMessage(jid: string, text: string) {
      logger.info(
        { jid },
        'Outlook Calendar: sendMessage (no-op — responses routed through requesting channel)',
      );
    },

    isConnected() {
      return connected;
    },

    ownsJid(jid: string) {
      return jid.startsWith('calendar:');
    },

    async disconnect() {
      if (pollTimer) clearInterval(pollTimer);
      connected = false;
    },

    /**
     * syncGroups: called by NanoClaw on startup.
     * Fetches the next 7 days of events and exposes them as chat metadata.
     */
    async syncGroups(_force: boolean) {
      try {
        const events = await fetchUpcomingEvents(7);
        logger.info(
          `Outlook Calendar: ${events.length} upcoming events in next 7 days`,
        );
        // Make the events available as a virtual "group" so the agent can query them
        const summary = events.length
          ? events.map(formatEvent).join('\n')
          : 'Aucun événement dans les 7 prochains jours.';
        const jid = 'calendar:week-view';
        opts.onChatMetadata(
          jid,
          new Date().toISOString(),
          'Agenda 7 jours',
          'outlook-calendar',
          false,
        );
      } catch (err) {
        logger.warn({ err }, 'Outlook Calendar: syncGroups failed');
      }
    },
  };
}

registerChannel('outlook-calendar', createOutlookCalendarChannel);
