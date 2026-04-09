/**
 * Veille professionnelle — channel RSS/Atom
 *
 * Agrège les flux RSS configurés et soumet les nouveaux articles à l'agent
 * Claude pour un résumé et une évaluation de pertinence.
 *
 * Required .env variables:
 *   NEWS_FEEDS       — URLs séparées par des virgules
 *                      e.g. https://feeds.lemonde.fr/rss/une,https://www.usine-digitale.fr/rss.xml
 *   NEWS_NOTIFY_JID  — JID WhatsApp pour livrer le digest (ex: 33612345678@s.whatsapp.net)
 *   NEWS_CRON        — Expression cron pour le digest (défaut: "0 8 * * *" = 8h chaque matin)
 *   NEWS_TOPICS      — Thèmes à surveiller, séparés par des virgules (optionnel)
 *                      e.g. "intelligence artificielle,droit urbanisme,collectivités locales"
 */

import RSSParser from 'rss-parser';
import { CronExpressionParser } from 'cron-parser';
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

const SEEN_PATH = path.join(STORE_DIR, 'news-seen.json');
const DEFAULT_CRON = '0 8 * * *'; // 8h00 chaque matin

// ── Seen-articles cache ────────────────────────────────────────────────────

function loadSeen(): Set<string> {
  try {
    const data = JSON.parse(fs.readFileSync(SEEN_PATH, 'utf-8'));
    return new Set(data);
  } catch {
    return new Set();
  }
}

function saveSeen(seen: Set<string>): void {
  fs.mkdirSync(path.dirname(SEEN_PATH), { recursive: true });
  // Keep only last 5000 entries to avoid unbounded growth
  const arr = [...seen].slice(-5000);
  fs.writeFileSync(SEEN_PATH, JSON.stringify(arr));
}

// ── Channel factory ────────────────────────────────────────────────────────

function createNewsFeedChannel(opts: {
  onMessage: OnInboundMessage;
  onChatMetadata: OnChatMetadata;
  registeredGroups: () => Record<string, import('../types.js').RegisteredGroup>;
}): Channel | null {
  const env = readEnvFile([
    'NEWS_FEEDS',
    'NEWS_NOTIFY_JID',
    'NEWS_CRON',
    'NEWS_TOPICS',
  ]);
  const feedUrls = (process.env.NEWS_FEEDS || env.NEWS_FEEDS || '')
    .split(',')
    .map((u) => u.trim())
    .filter(Boolean);
  const notifyJid = process.env.NEWS_NOTIFY_JID || env.NEWS_NOTIFY_JID || '';
  const cronExpr = process.env.NEWS_CRON || env.NEWS_CRON || DEFAULT_CRON;
  const topics = (process.env.NEWS_TOPICS || env.NEWS_TOPICS || '')
    .split(',')
    .map((t) => t.trim())
    .filter(Boolean);

  if (feedUrls.length === 0) {
    logger.warn('News Feed: NEWS_FEEDS not set — channel skipped');
    return null;
  }

  const parser = new RSSParser();
  let connected = false;
  let cronTimer: NodeJS.Timeout | null = null;
  let seen = loadSeen();

  async function fetchNewArticles(): Promise<
    Array<{
      title: string;
      link: string;
      summary: string;
      source: string;
      date: string;
    }>
  > {
    const results: Array<{
      title: string;
      link: string;
      summary: string;
      source: string;
      date: string;
    }> = [];

    for (const url of feedUrls) {
      try {
        const feed = await parser.parseURL(url);
        const sourceName = feed.title ?? url;
        for (const item of feed.items ?? []) {
          const id = item.guid || item.link || item.title || '';
          if (!id || seen.has(id)) continue;
          seen.add(id);
          results.push({
            title: item.title ?? '(sans titre)',
            link: item.link ?? '',
            summary: item.contentSnippet ?? item.content?.slice(0, 300) ?? '',
            source: sourceName,
            date: item.isoDate ?? new Date().toISOString(),
          });
        }
      } catch (err) {
        logger.warn({ err, url }, 'News Feed: failed to fetch feed');
      }
    }

    saveSeen(seen);
    return results;
  }

  type Article = {
    title: string;
    link: string;
    summary: string;
    source: string;
    date: string;
  };

  function buildDigestContent(articles: Article[]): string {
    if (articles.length === 0)
      return '🗞️ **Veille du jour** — Aucun nouvel article depuis le dernier digest.';

    const topicsLine = topics.length
      ? `\nThèmes surveillés : ${topics.join(', ')}\n`
      : '';
    const lines = articles
      .slice(0, 20)
      .map(
        (a) =>
          `• **${a.title}** (${a.source})\n  ${a.summary.slice(0, 120)}…\n  🔗 ${a.link}`,
      );

    const date = new Date().toLocaleDateString('fr-FR', { timeZone: TIMEZONE });
    return `🗞️ **Veille du jour — ${date}**${topicsLine}\n\n${lines.join('\n\n')}`;
  }

  async function runDigest(): Promise<void> {
    try {
      const articles = await fetchNewArticles();
      const content = buildDigestContent(articles);
      const jid = 'news:digest';

      const msg: NewMessage = {
        id: `news-digest-${Date.now()}`,
        chat_jid: jid,
        sender: 'news-feed',
        sender_name: 'Veille',
        content,
        timestamp: new Date().toISOString(),
        is_from_me: false,
      };

      opts.onChatMetadata(
        jid,
        new Date().toISOString(),
        'Veille professionnelle',
        'news-feed',
        false,
      );
      opts.onMessage(jid, msg);
      logger.info(`News Feed: digest sent (${articles.length} new articles)`);
    } catch (err) {
      logger.error({ err }, 'News Feed: digest failed');
    }
  }

  function scheduleNextRun(): void {
    try {
      const interval = CronExpressionParser.parse(cronExpr, { tz: TIMEZONE });
      const nextDate = interval.next().toDate();
      const delay = nextDate.getTime() - Date.now();
      logger.info(
        `News Feed: next digest at ${nextDate.toLocaleString('fr-FR', { timeZone: TIMEZONE })}`,
      );
      cronTimer = setTimeout(async () => {
        await runDigest();
        scheduleNextRun(); // reschedule
      }, delay);
    } catch (err) {
      logger.error(
        { err, cronExpr },
        'News Feed: invalid cron expression, using 24h interval',
      );
      cronTimer = setInterval(runDigest, 24 * 3600 * 1000);
    }
  }

  return {
    name: 'news-feed',

    async connect() {
      connected = true;
      logger.info(
        `News Feed: connected ✓ — watching ${feedUrls.length} feed(s), cron: ${cronExpr}`,
      );
      if (topics.length) logger.info(`News Feed: topics: ${topics.join(', ')}`);
      scheduleNextRun();
    },

    async sendMessage(_jid: string, _text: string) {
      // Veille is one-way (push only). Responses go through the requesting channel.
    },

    isConnected() {
      return connected;
    },

    ownsJid(jid: string) {
      return jid.startsWith('news:');
    },

    async disconnect() {
      if (cronTimer) clearTimeout(cronTimer);
      connected = false;
    },
  };
}

registerChannel('news-feed', createNewsFeedChannel);
