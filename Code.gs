/**
 * Googleカレンダーの「今日の予定」をNotion DBへUPSERT（既存更新 or 新規作成）
 * 必要なスクリプトプロパティ:
 *   NOTION_TOKEN   : Notion Internal Integration Token (secret_...)
 *   NOTION_DB_ID   : Notion Database ID
 *   TIMEZONE       : 例 "Asia/Tokyo"（省略可。未設定ならAsia/Tokyoを使用）
 *
 * Notion DB には以下プロパティを事前作成してください:
 *   Name(Title), Date(Date), EventId(Rich text), Calendar(Select),
 *   Location(Rich text), Attendees(Multi-select), EventURL(URL), Description(Rich text)
 */

const CONF = (() => {
  const props = PropertiesService.getScriptProperties();
  return {
    NOTION_TOKEN: props.getProperty('NOTION_TOKEN'),
    NOTION_DB_ID: props.getProperty('NOTION_DB_ID'),
    TIMEZONE: props.getProperty('TIMEZONE') || 'Asia/Tokyo',
    NOTION_API: 'https://api.notion.com/v1',
    NOTION_VERSION: '2022-06-28', // 安定版
  };
})();

/** エントリーポイント */
function syncTodayEventsToNotion() {
  assertConfig();

  const tz = CONF.TIMEZONE;
  const now = new Date();
  const today = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd 00:00:00'));
  const tomorrow = new Date(Utilities.formatDate(new Date(today.getTime() + 24*3600*1000), tz, 'yyyy/MM/dd 00:00:00'));

  // すべてのカレンダーから今日のイベントを取得
  const calendars = CalendarApp.getAllCalendars();
  let countUpserted = 0;

  calendars.forEach(cal => {
    const events = cal.getEvents(today, tomorrow);
    events.forEach(ev => {
      const payload = buildNotionPropertiesFromEvent(ev, cal.getName(), tz);
      const eventId = getEventStableId(ev);

      // EventIdで既存検索 → 更新 or 作成
      const existing = notionFindByEventId(eventId);
      if (existing && existing.length > 0) {
        const pageId = existing[0].id;
        notionUpdatePage(pageId, payload);
      } else {
        notionCreatePage(CONF.NOTION_DB_ID, payload);
      }
      countUpserted++;
    });
  });

  Logger.log(`Done. Upserted ${countUpserted} events.`);
}

/** ========= Helper: カレンダーイベント → Notionプロパティ ========= */
function buildNotionPropertiesFromEvent(event, calendarName, tz) {
  const title = event.getTitle() || '(無題の予定)';
  const isAllDay = event.isAllDayEvent();
  const start = event.getStartTime();
  const end = event.getEndTime();

  const dateProp = isAllDay
    ? { start: formatYMD(start, tz) } // 終日は開始日のみ（必要ならendも調整して入れる）
    : { start: toISOStringUTC(start), end: toISOStringUTC(end) };

  const attendees = (event.getGuestList() || []).map(g => g.getEmail()).slice(0, 50);
  const location = event.getLocation() || '';
  const description = (event.getDescription && event.getDescription()) || '';
  const url = (event.getHtmlLink && event.getHtmlLink()) || (event.getUrl && event.getUrl()) || '';

  const props = {
    // Title
    'Name': { title: [{ text: { content: title } }] },
    // Date
    'Date': { date: dateProp },
    // EventId（重複防止）
    'EventId': { rich_text: [{ text: { content: getEventStableId(event) } }] },
    // Calendar名（Select）: 未定義のオプションは自動で作る
    'Calendar': { select: { name: calendarName || 'Default' } },
    // Location
    'Location': location
      ? { rich_text: [{ text: { content: location } }] }
      : { rich_text: [] },
    // Attendees（Multi-select）
    'Attendees': { multi_select: attendees.map(a => ({ name: a })) },
    // URL
    'EventURL': url ? { url } : { url: null },
    // 説明
    'Description': description
      ? { rich_text: [{ text: { content: truncate(description, 1900) } }] }
      : { rich_text: [] },
  };

  return { properties: props };
}

/** できるだけ安定するID（Apps ScriptのgetIdでOK。@google.com を削る場合も） */
function getEventStableId(event) {
  let id = '';
  try { id = event.getId(); } catch (e) {}
  if (!id) {
    // フォールバック：開始時刻＋タイトルで簡易ID（重複の可能性は低いがゼロではない）
    id = `${event.getStartTime().getTime()}::${event.getTitle()}`;
  }
  // GoogleカレンダーのIDに「@google.com」サフィックスがつく場合は除去
  return id.replace(/@google\.com$/i, '');
}

/** ========= Helper: Notion API ========= */
function notionHeaders(json = true) {
  const h = {
    'Authorization': `Bearer ${CONF.NOTION_TOKEN}`,
    'Notion-Version': CONF.NOTION_VERSION,
  };
  if (json) h['Content-Type'] = 'application/json; charset=utf-8';
  return h;
}

function notionCreatePage(databaseId, payload) {
  const url = `${CONF.NOTION_API}/pages`;
  const body = {
    parent: { database_id: databaseId },
    ...payload,
  };
  return fetchJson(url, {
    method: 'post',
    headers: notionHeaders(true),
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
}

function notionUpdatePage(pageId, payload) {
  const url = `${CONF.NOTION_API}/pages/${pageId}`;
  return fetchJson(url, {
    method: 'patch',
    headers: notionHeaders(true),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
}

function notionFindByEventId(eventId) {
  const url = `${CONF.NOTION_API}/databases/${CONF.NOTION_DB_ID}/query`;
  const body = {
    filter: {
      property: 'EventId',
      rich_text: { equals: eventId },
    },
    page_size: 1,
  };
  const res = fetchJson(url, {
    method: 'post',
    headers: notionHeaders(true),
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  return (res && res.results) || [];
}

/** ========= Utilities ========= */
function toISOStringUTC(d) {
  // ISO8601（UTC, Z付き）— Notionが受け付ける形式
  return new Date(d).toISOString();
}

function formatYMD(d, tz) {
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function truncate(s, max) {
  if (!s) return s;
  return s.length <= max ? s : s.slice(0, max - 3) + '...';
}

function fetchJson(url, options) {
  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const text = resp.getContentText('utf-8');
  if (code >= 200 && code < 300) {
    return JSON.parse(text);
  }
  // 429/5xxなど簡易リトライ（1回だけ）
  if ((code === 429 || code >= 500) && !options.__retried) {
    Utilities.sleep(1000);
    return fetchJson(url, { ...options, __retried: true });
  }
  throw new Error(`HTTP ${code}: ${text}`);
}

function assertConfig() {
  if (!CONF.NOTION_TOKEN) throw new Error('NOTION_TOKEN (Script Property) が未設定です。');
  if (!CONF.NOTION_DB_ID) throw new Error('NOTION_DB_ID (Script Property) が未設定です。');
}

/***** ▼ 設定（必要なら変更） *************************************/
// 同期対象カレンダー（名前一致 or 部分一致）— 絞らないなら空配列のままでOK
const TARGET_CAL_NAMES = []; // 例: ['仕事', '個人']
// 月間ポーリングの頻度に合わせて「何日先までを見るか」
// 例: 月初フル同期 + 1時間/1日ごとの追従 → 今月の月末まで見れば十分
const LOOKAHEAD_DAYS = 40; // 当月末＋数日の余裕
/***** ▲ 設定 *****************************************************/

/** 月初フル同期（当月）→ 作成/更新 + Notion側だけに残ったものをアーカイブ */
function runMonthlySync() {
  assertConfig();
  const tz = CONF.TIMEZONE;
  const { start, end } = monthBounds(new Date(), tz); // 当月
  doUpsertForRange(start, end);
  archiveNotInCalendarForRange(start, end);
}

/** 定期ポーリング（増減追従：新規/変更/削除）— 15分/1時間/1日などお好みで */
function runRollingSync() {
  assertConfig();
  const tz = CONF.TIMEZONE;
  const now = new Date();
  const start = firstDayOfMonth(now, tz);
  const end   = new Date(now.getTime() + LOOKAHEAD_DAYS * 24*3600*1000);
  doUpsertForRange(start, end);
  archiveNotInCalendarForRange(start, end);
}

/** 範囲のUPSERT（作成/更新） */
function doUpsertForRange(start, end) {
  const tz = CONF.TIMEZONE;
  const calendars = getTargetCalendars();
  const existingPropNames = notionGetDatabasePropNames();

  calendars.forEach(cal => {
    cal.getEvents(start, end).forEach(ev => {
      const rawPayload = buildNotionPropertiesFromEvent(ev, cal.getName(), tz);
      const payload = prunePropertiesToExisting(rawPayload, existingPropNames);
      const eventId = getEventStableId(ev);
      const existing = notionFindByEventId(eventId);
      if (existing && existing.length > 0) {
        notionUpdatePage(existing[0].id, payload);
        Logger.log(`updated: ${eventId} "${ev.getTitle()}"`);
      } else {
        notionCreatePage(CONF.NOTION_DB_ID, payload);
        Logger.log(`created: ${eventId} "${ev.getTitle()}"`);
      }
    });
  });
}

/** 範囲で「カレンダーには無いEventId」をNotion側でアーカイブ */
function archiveNotInCalendarForRange(start, end) {
  const datePropName = (typeof PROPS !== 'undefined' && PROPS.date) ? PROPS.date : 'Date';
  const calendars = getTargetCalendars();
  const calIds = new Set();

  calendars.forEach(cal => {
    cal.getEvents(start, end).forEach(ev => calIds.add(getEventStableId(ev)));
  });

  const notionPages = notionQueryByDateRangeAll(PROPS.date, start, end);
  let archived = 0;
  notionPages.forEach(p => {
    const eid = getRichTextPlain(p.properties?.[PROPS.eventId]);
    if (eid && !calIds.has(eid)) {
      notionArchivePage(p.id);
      archived++;
    }
  });
  Logger.log(`archived (not in Google Calendar): ${archived}`);
}

/***** ▼ サポート関数 *********************************************/

/** 対象カレンダーの抽出（未指定なら全カレンダー） */
function getTargetCalendars() {
  const all = CalendarApp.getAllCalendars();
  if (!TARGET_CAL_NAMES || TARGET_CAL_NAMES.length === 0) return all;
  return all.filter(cal => {
    const name = cal.getName() || '';
    return TARGET_CAL_NAMES.some(t => name.indexOf(t) >= 0);
  });
}

/** 当月の開始/終了（[月初00:00, 翌月初00:00)） */
function monthBounds(baseDate, tz) {
  const s = firstDayOfMonth(baseDate, tz);
  const e = firstDayOfNextMonth(baseDate, tz);
  return { start: s, end: e };
}
function firstDayOfMonth(d, tz) {
  const y = Number(Utilities.formatDate(d, tz, 'yyyy'));
  const m = Number(Utilities.formatDate(d, tz, 'MM')) - 1;
  return new Date(y, m, 1, 0, 0, 0);
}
function firstDayOfNextMonth(d, tz) {
  const y = Number(Utilities.formatDate(d, tz, 'yyyy'));
  const m = Number(Utilities.formatDate(d, tz, 'MM')) - 1;
  return new Date(y, m + 1, 1, 0, 0, 0);
}

/** Notion: 範囲に一致するページを（ページネーションで）すべて取得 */
function notionQueryByDateRangeAll(datePropName, startDate, endDate) {
  const results = [];
  let has_more = true, cursor = null;

  while (has_more) {
    const body = {
      filter: {
        and: [
          { property: datePropName, date: { on_or_after: toISOStringUTC(startDate) } },
          { property: datePropName, date: { on_or_before: toISOStringUTC(endDate) } },
        ]
      },
      page_size: 100
    };
    if (cursor) body.start_cursor = cursor;

    const res = fetchJson(`${CONF.NOTION_API}/databases/${CONF.NOTION_DB_ID}/query`, {
      method: 'post',
      headers: notionHeaders(true),
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });

    (res.results || []).forEach(r => results.push(r));
    has_more = !!res.has_more;
    cursor = res.next_cursor || null;
  }
  return results;
}

/** Notion: ページをアーカイブ（物理削除せず安全） */
function notionArchivePage(pageId) {
  return fetchJson(`${CONF.NOTION_API}/pages/${pageId}`, {
    method: 'patch',
    headers: notionHeaders(true),
    payload: JSON.stringify({ archived: true }),
    muteHttpExceptions: true,
  });
}

/** Rich textのプレーン文字を取り出し */
function getRichTextPlain(richTextProp) {
  try {
    const arr = richTextProp.rich_text || [];
    return arr.map(x => x.plain_text || x.text?.content || '').join('');
  } catch (e) { return ''; }
}
// DBに存在するプロパティ名の集合を取得
function notionGetDatabasePropNames() {
  const url = `${CONF.NOTION_API}/databases/${CONF.NOTION_DB_ID}`;
  const res = fetchJson(url, {
    method: 'get',
    headers: notionHeaders(false),
    muteHttpExceptions: true,
  });
  const names = new Set(Object.keys(res.properties || {}));
  return names;
}

// payload.properties から、DBに無い列を削る
function prunePropertiesToExisting(payload, existingNames) {
  const out = {};
  for (const [k, v] of Object.entries(payload.properties || {})) {
    if (existingNames.has(k)) out[k] = v;
  }
  return { properties: out };
}
// Notion列名マッピング（あなたのDBの実名に合わせる）
const PROPS = {
  title: 'Name',        // タイトル
  date: 'Date',         // ← Date（日本語なら '日付'）
  eventId: 'EventId',   // ← EventId（大文字・小文字注意）
  calendar: 'Calendar',
  location: 'Location',
  attendees: 'Attendees',
  url: 'EventURL',
  description: 'Description',
};
