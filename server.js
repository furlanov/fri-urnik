const crypto = require("node:crypto");
const fs = require("node:fs/promises");
const http = require("node:http");
const path = require("node:path");

const PORT = Number.parseInt(process.env.PORT || "8080", 10);
const PUBLIC_DIR = path.join(__dirname, "public");
const DEFAULT_TIMETABLE_URL = process.env.ENRICH_URL || process.env.TIMETABLE_URL || "";
const DATA_DIR = path.resolve(process.env.DATA_DIR || path.join(__dirname, "data"));
const IMPORT_CACHE_PATH = path.join(DATA_DIR, "import-cache.json");
const HOLIDAY_CACHE_PATH = path.join(DATA_DIR, "holiday-cache.json");
const CONFIG_PATH = path.join(DATA_DIR, "config.json");
const TIMEZONE = process.env.TZ || "Europe/Ljubljana";
const HOLIDAY_API_BASE_URL = "https://date.nager.at/api/v3/PublicHolidays";

const DAY_CODES = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];
const TYPE_LABELS = {
  P: { label: "Lecture", role: "lecture" },
  LV: { label: "Lab", role: "lab" },
  AV: { label: "Tutorial", role: "tutorial" },
};
const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
};

const scheduleCache = new Map();
let importStorePromise = null;
let holidayStorePromise = null;
let configStorePromise = null;

function hashId(value) {
  return crypto.createHash("sha1").update(value).digest("hex").slice(0, 12);
}

function sendJson(res, status, payload, extraHeaders = {}) {
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "cache-control": "no-store",
    ...securityHeaders(),
    ...extraHeaders,
  });
  res.end(JSON.stringify(payload));
}

function sendText(res, status, content, headers = {}) {
  res.writeHead(status, {
    "cache-control": "no-store",
    ...securityHeaders(),
    ...headers,
  });
  res.end(content);
}

function securityHeaders() {
  return {
    "content-security-policy":
      "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; base-uri 'none'; form-action 'none'",
    "referrer-policy": "same-origin",
    "x-content-type-options": "nosniff",
    "x-frame-options": "DENY",
  };
}

function httpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

async function getSchedule(options = {}) {
  const importData = await getTimetableImport(options.sourceUrl, {
    force: options.forceImport,
  });
  if (!importData.icalContent) {
    throw httpError(502, "Timetable import did not return calendar data");
  }

  const cached = scheduleCache.get(importData.url);
  let payload = cached?.sourceVersion === importData.version ? cached.payload : null;

  if (!payload) {
    const ical = importData.icalContent;
    payload = parseCalendar(
      ical,
      {
        sourceName: importData.sourceName || "FRI timetable",
        sourceUrl: importData.url,
        modified: importData.fetchedAt ? new Date(importData.fetchedAt) : new Date(),
        size: Buffer.byteLength(ical, "utf8"),
      },
      importData
    );

    scheduleCache.set(importData.url, {
      sourceVersion: importData.version,
      payload,
    });
  }

  const holidays = await getHolidayCalendar(payload.range, {
    force: options.forceImport,
  });

  return {
    ...payload,
    holidays,
  };
}

function emptyImport(error = "") {
  return {
    enabled: false,
    url: DEFAULT_TIMETABLE_URL,
    icalUrl: "",
    fetchedAt: null,
    dayKey: "",
    version: "none",
    sourceName: "",
    icalContent: "",
    localMtimeMs: null,
    map: new Map(),
    context: { title: "", subtitle: "" },
    error,
  };
}

async function getTimetableImport(rawUrl, options = {}) {
  const sourceUrl = normalizeTimetableUrl(rawUrl || DEFAULT_TIMETABLE_URL);
  if (!sourceUrl) {
    throw httpError(400, "Timetable link is required");
  }

  const store = await loadImportStore();
  const key = hashId(sourceUrl);
  const cached = store.imports[key];
  const today = localDayKey(new Date());

  if (cached && cached.dayKey === today && !options.force) {
    return importFromCache(cached, "");
  }

  try {
    const htmlResponse = await fetch(sourceUrl, {
      headers: {
        accept: "text/html,application/xhtml+xml",
        "user-agent": "urnik-viewer/1.0",
      },
      signal: AbortSignal.timeout(10000),
    });

    if (!htmlResponse.ok) {
      throw new Error(`Timetable import returned ${htmlResponse.status}`);
    }

    const html = await htmlResponse.text();
    const map = parseAllocationEnrichment(html);
    const context = parsePageContext(html);
    const icalUrl = normalizeIcalUrl(extractIcalUrl(html, sourceUrl));
    const icalResponse = await fetch(icalUrl, {
      headers: {
        accept: "text/calendar,text/plain",
        "user-agent": "urnik-viewer/1.0",
      },
      signal: AbortSignal.timeout(10000),
    });

    if (!icalResponse.ok) {
      throw new Error(`Calendar import returned ${icalResponse.status}`);
    }

    const icalContent = await icalResponse.text();
    const entries = [...map.values()];
    const htmlHash = hashId(
      entries
        .map((value) => `${value.allocationId}:${value.courseName}:${value.typeCode}`)
        .sort()
        .join("|")
    );
    const icalHash = hashId(icalContent);
    const previousVersion = cached?.version || "";
    const version = hashId(`${htmlHash}:${icalHash}`);
    const record = {
      url: sourceUrl,
      icalUrl,
      dayKey: today,
      fetchedAt: new Date().toISOString(),
      htmlHash,
      icalHash,
      version,
      changed: previousVersion ? previousVersion !== version : true,
      entries,
      context,
      icalContent,
      error: "",
    };

    store.imports[key] = record;
    await saveImportStore(store);

    return importFromCache(record, "");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (cached?.icalContent) {
      console.error(`Using cached timetable import: ${message}`);
      return importFromCache(cached, message);
    }

    console.error(`Timetable import unavailable: ${message}`);
    throw httpError(502, `Timetable import unavailable: ${message}`);
  }
}

function importFromCache(record, error) {
  const map = new Map(
    (record.entries || []).map((entry) => [String(entry.allocationId), entry])
  );

  return {
    enabled: true,
    url: record.url,
    icalUrl: record.icalUrl,
    fetchedAt: record.fetchedAt,
    dayKey: record.dayKey,
    version: record.version,
    sourceName: "FRI timetable import",
    icalContent: record.icalContent,
    localMtimeMs: null,
    map,
    context: record.context || { title: "", subtitle: "" },
    changed: Boolean(record.changed),
    error,
  };
}

async function loadImportStore() {
  if (!importStorePromise) {
    importStorePromise = fs
      .readFile(IMPORT_CACHE_PATH, "utf8")
      .then((content) => JSON.parse(content))
      .catch((error) => {
        if (error.code === "ENOENT") return { imports: {} };
        throw error;
      });
  }
  const store = await importStorePromise;
  store.imports ||= {};
  return store;
}

async function saveImportStore(store) {
  await writeJsonAtomic(IMPORT_CACHE_PATH, store);
}

async function loadConfigStore() {
  if (!configStorePromise) {
    configStorePromise = fs
      .readFile(CONFIG_PATH, "utf8")
      .then((content) => JSON.parse(content))
      .catch((error) => {
        if (error.code === "ENOENT") return { sourceUrl: "", filters: {}, theme: "" };
        throw error;
      });
  }
  const store = await configStorePromise;
  store.filters ||= {};
  if (typeof store.sourceUrl !== "string") store.sourceUrl = "";
  if (!["light", "dark"].includes(store.theme)) store.theme = "";
  return store;
}

async function saveConfigStore(store) {
  configStorePromise = Promise.resolve(store);
  await writeJsonAtomic(CONFIG_PATH, store);
}

async function loadHolidayStore() {
  if (!holidayStorePromise) {
    holidayStorePromise = fs
      .readFile(HOLIDAY_CACHE_PATH, "utf8")
      .then((content) => JSON.parse(content))
      .catch((error) => {
        if (error.code === "ENOENT") return { years: {} };
        throw error;
      });
  }
  const store = await holidayStorePromise;
  store.years ||= {};
  return store;
}

async function saveHolidayStore(store) {
  holidayStorePromise = Promise.resolve(store);
  await writeJsonAtomic(HOLIDAY_CACHE_PATH, store);
}

async function writeJsonAtomic(destination, payload) {
  await fs.mkdir(DATA_DIR, { recursive: true });
  const tempPath = `${destination}.${process.pid}.${Date.now()}.${Math.random()
    .toString(16)
    .slice(2)}.tmp`;
  await fs.writeFile(tempPath, JSON.stringify(payload, null, 2));
  await fs.rename(tempPath, destination);
}

async function getHolidayCalendar(range, options = {}) {
  const years = holidayYearsForRange(range);
  if (!years.length) return { dates: {} };

  const store = await loadHolidayStore();
  const dates = {};
  let changed = false;

  for (const year of years) {
    const existing = store.years[String(year)];
    let holidays = existing?.holidays;

    if (!holidays || options.force) {
      const fetched = await fetchHolidayYear(year);
      if (fetched) {
        store.years[String(year)] = { year, fetchedAt: new Date().toISOString(), holidays: fetched };
        holidays = fetched;
        changed = true;
      }
    }

    for (const holiday of holidays || []) dates[holiday.date] = holiday;
  }

  if (changed) await saveHolidayStore(store);
  return { dates };
}

async function fetchHolidayYear(year) {
  try {
    const response = await fetch(`${HOLIDAY_API_BASE_URL}/${year}/SI`, {
      headers: { accept: "application/json", "user-agent": "urnik-viewer/1.0" },
      signal: AbortSignal.timeout(10000),
    });
    if (!response.ok) throw new Error(`Holiday API returned ${response.status}`);
    return normalizeHolidayPayload(await response.json());
  } catch (error) {
    console.error(`Holiday fetch failed for ${year}: ${error.message || error}`);
    return null;
  }
}

function holidayYearsForRange(range) {
  if (!range?.start || !range?.end) return [];
  const start = parseDateKey(range.start);
  const end = parseDateKey(range.end);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return [];
  const years = new Set();
  for (let year = start.getFullYear(); year <= end.getFullYear(); year += 1) {
    years.add(year);
  }
  return [...years];
}

function normalizeHolidayPayload(payload) {
  const byDate = new Map();
  for (const entry of Array.isArray(payload) ? payload : []) {
    if (!entry || entry.global !== true) continue;
    if (!Array.isArray(entry.types) || !entry.types.includes("Public")) continue;
    if (typeof entry.date !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(entry.date)) continue;
    const name = String(entry.localName || entry.name || "").trim() || "Dela prost dan";
    byDate.set(entry.date, { date: entry.date, name });
  }
  return [...byDate.values()].sort((a, b) => a.date.localeCompare(b.date));
}

async function readJsonBody(req, maxBytes = 128 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    req.on("data", (chunk) => {
      total += chunk.length;
      if (total > maxBytes) {
        req.destroy();
        reject(httpError(413, "Payload too large"));
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => {
      if (total === 0) return resolve({});
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString("utf8")));
      } catch {
        reject(httpError(400, "Invalid JSON body"));
      }
    });
    req.on("error", reject);
  });
}

function sanitizeFilterMap(input) {
  if (!input || typeof input !== "object") return {};
  const output = {};
  let keys = 0;
  for (const [rawKey, rawValue] of Object.entries(input)) {
    if (keys >= 2000) break;
    if (typeof rawKey !== "string" || rawKey.length > 64) continue;
    if (!/^[A-Za-z0-9_:.-]+$/.test(rawKey)) continue;
    output[rawKey] = Boolean(rawValue);
    keys += 1;
  }
  return output;
}

function sanitizeFilterStore(input) {
  if (!input || typeof input !== "object") return {};
  const output = {};
  let keys = 0;
  for (const [rawKey, rawValue] of Object.entries(input)) {
    if (keys >= 32) break;
    if (typeof rawKey !== "string" || rawKey.length > 512) continue;
    output[rawKey] = sanitizeFilterMap(rawValue);
    keys += 1;
  }
  return output;
}

function normalizeTimetableUrl(rawUrl) {
  const trimmed = String(rawUrl || "").trim();
  if (!trimmed) return "";

  let parsed;
  try {
    parsed = new URL(trimmed);
  } catch {
    throw httpError(400, "Invalid timetable URL");
  }

  if (
    parsed.protocol !== "https:" ||
    parsed.hostname !== "urnik.fri.uni-lj.si" ||
    !/^\/timetable\/[^/]+\/allocations$/.test(parsed.pathname)
  ) {
    throw httpError(400, "Only FRI timetable allocation URLs are allowed");
  }

  parsed.hash = "";
  parsed.search = parsed.searchParams.toString();
  return parsed.toString();
}

function normalizeIcalUrl(rawUrl) {
  const parsed = new URL(rawUrl);
  if (
    parsed.protocol !== "https:" ||
    parsed.hostname !== "urnik.fri.uni-lj.si" ||
    !/^\/timetable\/[^/]+\/allocations_ical$/.test(parsed.pathname)
  ) {
    throw new Error("Invalid calendar export URL in timetable page");
  }
  parsed.hash = "";
  parsed.search = parsed.searchParams.toString();
  return parsed.toString();
}

function extractIcalUrl(html, sourceUrl) {
  const href =
    html.match(/href="([^"]*allocations_ical[^"]*)"/)?.[1] ||
    html.match(/href='([^']*allocations_ical[^']*)'/)?.[1] ||
    "";
  if (href) {
    return new URL(decodeHtml(href), sourceUrl).toString();
  }

  const parsed = new URL(sourceUrl);
  parsed.pathname = parsed.pathname.replace(/\/allocations$/, "/allocations_ical");
  return parsed.toString();
}

function localDayKey(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function unfoldLines(content) {
  const lines = content.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const unfolded = [];

  for (const line of lines) {
    if ((line.startsWith(" ") || line.startsWith("\t")) && unfolded.length) {
      unfolded[unfolded.length - 1] += line.slice(1);
    } else {
      unfolded.push(line);
    }
  }

  return unfolded;
}

function parseCalendar(content, source, enrichment = emptyImport()) {
  const lines = unfoldLines(content);
  const calendarProps = [];
  const rawEvents = [];
  let current = null;

  for (const line of lines) {
    if (line === "BEGIN:VEVENT") {
      current = [];
    } else if (line === "END:VEVENT") {
      if (current) {
        rawEvents.push(current);
      }
      current = null;
    } else if (current) {
      current.push(line);
    } else if (line && !line.startsWith("BEGIN:") && !line.startsWith("END:")) {
      calendarProps.push(line);
    }
  }

  const events = rawEvents
    .map((eventLines, index) => parseEvent(eventLines, index, enrichment.map))
    .filter(Boolean)
    .sort((a, b) => {
      if (a.weekday !== b.weekday) return a.weekday - b.weekday;
      if (a.startMinute !== b.startMinute) return a.startMinute - b.startMinute;
      return a.displaySubject.localeCompare(b.displaySubject);
    });

  return {
    generatedAt: new Date().toISOString(),
    source: {
      name: source.sourceName,
      file: source.sourceName,
      url: source.sourceUrl,
      modified: source.modified.toISOString(),
      size: source.size,
    },
    context: enrichment.context || { title: "", subtitle: "" },
    calendar: {
      productId: getCalendarProperty(calendarProps, "PRODID") || "",
      eventCount: events.length,
      subjectNamesAvailable: events.some((event) => !event.subjectMissing),
    },
    enrichment: {
      enabled: enrichment.enabled,
      url: enrichment.url,
      icalUrl: enrichment.icalUrl,
      fetchedAt: enrichment.fetchedAt,
      dayKey: enrichment.dayKey,
      available: enrichment.map.size,
      applied: events.filter((event) => event.enriched).length,
      changed: Boolean(enrichment.changed),
      source: enrichment.error ? "cache" : "remote",
      error: enrichment.error,
    },
    range: buildRange(events),
    filters: buildFilters(events),
    events,
  };
}

function getCalendarProperty(lines, name) {
  const prefix = `${name}:`;
  const line = lines.find((item) => item.toUpperCase().startsWith(prefix));
  return line ? decodeIcalText(line.slice(prefix.length)) : "";
}

function parseEvent(lines, index, enrichmentMap = new Map()) {
  const props = toPropertyMap(lines);
  const summary = getFirstValue(props, "SUMMARY") || "Untitled";
  const description = getFirstValue(props, "DESCRIPTION") || "";
  const location = getFirstValue(props, "LOCATION") || "";
  const uid = getFirstValue(props, "UID") || `event-${index}`;
  const dtStartProp = getFirstProp(props, "DTSTART");
  const dtEndProp = getFirstProp(props, "DTEND");

  if (!dtStartProp || !dtEndProp) {
    return null;
  }

  const startDate = parseIcalDate(dtStartProp.rawValue);
  const endDate = parseIcalDate(dtEndProp.rawValue);
  if (!startDate || !endDate) {
    return null;
  }

  const summaryParts = parseSummary(summary);
  const descriptionParts = parseDescription(description);
  const allocationId = parseAllocationId(uid);
  const enrichment = allocationId ? enrichmentMap.get(allocationId) : null;
  const rawSubject = summaryParts.subject || descriptionParts.subject || "Unknown subject";
  const enrichedSubject = enrichment?.courseName || "";
  const subject =
    isMissingSubject(rawSubject) && enrichedSubject ? enrichedSubject : rawSubject;
  const typeCode =
    summaryParts.typeCode || descriptionParts.typeCode || enrichment?.typeCode || "OTHER";
  const type = TYPE_LABELS[typeCode] || { label: typeCode, role: "other" };
  const enrichedTeachers = (enrichment?.teachers || []).filter(Boolean);
  const teacher = enrichedTeachers.length
    ? enrichedTeachers.join(", ")
    : descriptionParts.teacher || "";
  const subjectMissing = isMissingSubject(subject);
  const displaySubject = displaySubjectName(subject);
  const filterLabel = subjectMissing
    ? displayInstructorName(teacher, location)
    : displaySubject;
  const filterKind = subjectMissing ? "instructor" : "subject";
  const rrule = getFirstValue(props, "RRULE");
  const recurrence = rrule ? parseRRule(rrule, startDate) : null;
  const startMinute = startDate.getHours() * 60 + startDate.getMinutes();
  const endMinute = endDate.getHours() * 60 + endDate.getMinutes();
  const durationMinutes = Math.max(15, Math.round((endDate - startDate) / 60000));
  const filterKey = hashId(`${filterKind}|${filterLabel}|${typeCode}`);
  const colorKey = hashId(filterLabel);

  return {
    id: hashId(`${uid}|${startDate.toISOString()}|${summary}`),
    uid,
    allocationId,
    summary,
    subject,
    rawSubject,
    displaySubject,
    subjectMissing,
    subjectSource: enrichment?.courseName ? "timetable" : "calendar",
    enriched: Boolean(enrichment?.courseName),
    activityName: enrichment?.activityName || "",
    activityShortName: enrichment?.shortName || "",
    courseCode: enrichment?.courseCode || "",
    groupLabels: enrichment?.groups || [],
    groupDescriptions: enrichment?.groupDescriptions || [],
    capacity: enrichment?.capacity || "",
    sourceColor: enrichment?.sourceColor || "",
    teacher,
    filterLabel,
    filterKind,
    location,
    descriptionLines: description
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean),
    type: {
      code: typeCode,
      label: type.label,
      role: type.role,
    },
    start: toLocalIso(startDate),
    end: toLocalIso(endDate),
    firstDate: toDateKey(startDate),
    lastDate: recurrence?.until ? toDateKey(recurrence.until) : toDateKey(startDate),
    weekday: startDate.getDay(),
    weekdayCode: DAY_CODES[startDate.getDay()],
    startMinute,
    endMinute,
    durationMinutes,
    recurrence: recurrence
      ? {
          frequency: recurrence.frequency,
          interval: recurrence.interval,
          byDay: recurrence.byDay,
          until: recurrence.until ? toLocalIso(recurrence.until) : null,
        }
      : null,
    filterKey,
    colorKey,
  };
}

function toPropertyMap(lines) {
  const map = new Map();

  for (const line of lines) {
    const prop = parseProperty(line);
    if (!prop) continue;

    if (!map.has(prop.name)) {
      map.set(prop.name, []);
    }
    map.get(prop.name).push(prop);
  }

  return map;
}

function parseProperty(line) {
  const colon = line.indexOf(":");
  if (colon === -1) return null;

  const left = line.slice(0, colon);
  const rawValue = line.slice(colon + 1);
  const [namePart, ...paramParts] = left.split(";");
  const params = {};

  for (const param of paramParts) {
    const equal = param.indexOf("=");
    if (equal === -1) continue;
    params[param.slice(0, equal).toUpperCase()] = param.slice(equal + 1);
  }

  return {
    name: namePart.toUpperCase(),
    params,
    rawValue,
    value: decodeIcalText(rawValue),
  };
}

function getFirstProp(props, name) {
  return props.get(name)?.[0] || null;
}

function getFirstValue(props, name) {
  return getFirstProp(props, name)?.value || "";
}

function decodeIcalText(value) {
  return value
    .replace(/\\[nN]/g, "\n")
    .replace(/\\,/g, ",")
    .replace(/\\;/g, ";")
    .replace(/\\\\/g, "\\");
}

function parseAllocationId(uid) {
  return uid.match(/urnikfri-(\d+)/)?.[1] || "";
}

function parseAllocationEnrichment(html) {
  const starts = [...html.matchAll(/<div class="grid-entry"/g)].map(
    (match) => match.index
  );
  const endIndex = html.search(/<\/body>/i);
  const map = new Map();

  for (let index = 0; index < starts.length; index += 1) {
    const block = html.slice(
      starts[index],
      starts[index + 1] || (endIndex === -1 ? html.length : endIndex)
    );
    const allocationId = getHtmlAttribute(block, "data-allocation-id");
    if (!allocationId) continue;

    const hover = block.match(/<div class="entry-hover">([\s\S]*?)<\/div>/)?.[1] || "";
    const hoverLines = hover
      .split(/<br\s*\/?>/i)
      .map(htmlText)
      .filter(Boolean)
      .filter((line) => !line.startsWith("<!--"));
    const shortName = extractClassText(block, "link-subject");
    const activityName =
      hoverLines.find((line) => parseActivityName(line).courseName) || shortName;
    const activity = parseActivityName(activityName);
    const typeFromPage = htmlText(
      block.match(/<span class="entry-type">([\s\S]*?)<\/span>/)?.[1] || ""
    )
      .replace("|", "")
      .trim()
      .toUpperCase();
    const capacity = hover.match(/<!--\s*\(\s*([^)]+?)\s*\)\s*-->/)?.[1].trim() || "";
    const sourceColor = extractSourceColor(block);
    const groupDescriptions = hoverLines.filter(isGroupDescription);

    map.set(allocationId, {
      allocationId,
      shortName,
      activityName: htmlText(activityName),
      courseName: activity.courseName,
      courseCode: activity.courseCode,
      typeCode: activity.typeCode || typeFromPage,
      day: getHtmlAttribute(block, "data-day"),
      start: getHtmlAttribute(block, "data-start"),
      durationHours: Number.parseInt(getHtmlAttribute(block, "data-duration"), 10) || null,
      classroom: extractClassText(block, "link-classroom"),
      teachers: extractAllClassText(block, "link-teacher"),
      groups: extractAllClassText(block, "link-group"),
      groupDescriptions,
      capacity,
      sourceColor,
    });
  }

  return map;
}

function parsePageContext(html) {
  const titles = html.match(/<div class="titles">([\s\S]*?)<\/div>\s*<div class="aside">/);
  const scope = titles ? titles[1] : html;
  const title = htmlText(scope.match(/<span class="title">([\s\S]*?)<\/span>/)?.[1] || "");
  const subtitle = htmlText(scope.match(/<span class="subtitle">([\s\S]*?)<\/span>/)?.[1] || "");
  return { title, subtitle };
}

function extractSourceColor(block) {
  const style = block.match(/style="([^"]*)"/)?.[1] || "";
  const color = style.match(/background-color:\s*([^;"]+)/i)?.[1].trim() || "";
  if (!color) return "";
  const hsla = color.match(/hsla?\(\s*([\d.]+)\s*,\s*([\d.]+)%\s*,\s*([\d.]+)%(?:\s*,\s*([\d.]+))?\s*\)/i);
  if (!hsla) return color;
  const [, h, s, l] = hsla;
  return `hsla(${Number(h).toFixed(0)}, ${Math.min(60, Number(s)).toFixed(0)}%, ${Math.max(72, Number(l)).toFixed(0)}%, 0.78)`;
}

function isGroupDescription(line) {
  return /letnik/i.test(line) && /skupina/i.test(line);
}

function parseActivityName(value) {
  const clean = htmlText(value);
  const fullMatch = clean.match(/^(.+?)\(([^)]+)\)_([A-Za-z0-9]+)$/);
  if (fullMatch) {
    return {
      courseName: fullMatch[1].trim(),
      courseCode: fullMatch[2].trim(),
      typeCode: fullMatch[3].trim().toUpperCase(),
    };
  }

  const shortMatch = clean.match(/^(.+?)_([A-Za-z0-9]+)$/);
  if (shortMatch && shortMatch[1].trim() && !/^[A-Z0-9()]+$/.test(shortMatch[1])) {
    return {
      courseName: shortMatch[1].trim(),
      courseCode: "",
      typeCode: shortMatch[2].trim().toUpperCase(),
    };
  }

  return {
    courseName: "",
    courseCode: "",
    typeCode: shortMatch?.[2]?.trim().toUpperCase() || "",
  };
}

function getHtmlAttribute(html, attribute) {
  const pattern = new RegExp(`${attribute}="([^"]*)"`);
  return decodeHtml(pattern.exec(html)?.[1] || "");
}

function extractClassText(html, className) {
  return (
    htmlText(
      new RegExp(`<a class="${className}"[^>]*>([\\s\\S]*?)<\\/a>`).exec(html)?.[1] ||
        ""
    ) || ""
  );
}

function extractAllClassText(html, className) {
  const pattern = new RegExp(`<a class="${className}"[^>]*>([\\s\\S]*?)<\\/a>`, "g");
  return [...html.matchAll(pattern)].map((match) => htmlText(match[1])).filter(Boolean);
}

function htmlText(value) {
  return decodeHtml(value.replace(/<[^>]*>/g, " ")).replace(/\s+/g, " ").trim();
}

function decodeHtml(value) {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) =>
      String.fromCodePoint(Number.parseInt(hex, 16))
    )
    .replace(/&#(\d+);/g, (_, number) =>
      String.fromCodePoint(Number.parseInt(number, 10))
    );
}

function parseIcalDate(value) {
  const trimmed = value.trim();
  const isUtc = trimmed.endsWith("Z");
  const clean = isUtc ? trimmed.slice(0, -1) : trimmed;
  const match = clean.match(
    /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?)?$/
  );
  if (!match) return null;

  const [, year, month, day, hour = "0", minute = "0", second = "0"] = match;
  const parts = [
    Number.parseInt(year, 10),
    Number.parseInt(month, 10) - 1,
    Number.parseInt(day, 10),
    Number.parseInt(hour, 10),
    Number.parseInt(minute, 10),
    Number.parseInt(second, 10),
  ];

  return isUtc ? new Date(Date.UTC(...parts)) : new Date(...parts);
}

function parseLocalDateTime(value) {
  const [datePart, timePart = "00:00:00"] = value.split("T");
  const [year, month, day] = datePart.split("-").map(Number);
  const [hour, minute, second = 0] = timePart.split(":").map(Number);
  return new Date(year, month - 1, day, hour, minute, second);
}

function parseSummary(summary) {
  const match = summary.match(/^(.*?)\s+-\s+([A-Za-z0-9_]+)$/);
  if (!match) {
    return { subject: summary.trim(), typeCode: "" };
  }

  return {
    subject: match[1].trim(),
    typeCode: match[2].trim().toUpperCase(),
  };
}

function parseDescription(description) {
  const lines = description
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  const firstLine = lines[0] || "";
  const match = firstLine.match(/^(.*?)\s+([A-Za-z0-9_]+)\s+@\s+(.+)$/);
  const teacher = lines.slice(1).join(", ");

  return {
    subject: match ? match[1].trim() : "",
    typeCode: match ? match[2].trim().toUpperCase() : "",
    room: match ? match[3].trim() : "",
    teacher,
  };
}

function isMissingSubject(subject) {
  return !subject || /^unknown subject$/i.test(subject.trim());
}

function displaySubjectName(subject) {
  if (isMissingSubject(subject)) {
    return "Subject not in export";
  }
  return subject.trim();
}

function displayInstructorName(teacher, location) {
  const primaryTeacher = teacher.split(",")[0].trim();
  if (primaryTeacher) return primaryTeacher;
  if (location) return `Room ${location}`;
  return "Unassigned";
}

function parseRRule(value, startDate) {
  const fields = Object.fromEntries(
    value.split(";").map((part) => {
      const [key, val = ""] = part.split("=");
      return [key.toUpperCase(), val];
    })
  );

  return {
    frequency: fields.FREQ || "",
    interval: Number.parseInt(fields.INTERVAL || "1", 10) || 1,
    byDay: fields.BYDAY
      ? fields.BYDAY.split(",").map((item) => item.trim()).filter(Boolean)
      : [DAY_CODES[startDate.getDay()]],
    until: fields.UNTIL ? parseIcalDate(fields.UNTIL) : null,
  };
}

function buildRange(events) {
  if (!events.length) {
    return {
      start: null,
      end: null,
      startHour: 7,
      endHour: 18,
    };
  }

  const startDates = events.map((event) => event.firstDate).sort();
  const endDates = events.map((event) => event.lastDate).sort();
  const minMinute = Math.min(...events.map((event) => event.startMinute));
  const maxMinute = Math.max(...events.map((event) => event.endMinute));

  return {
    start: startDates[0],
    end: endDates[endDates.length - 1],
    startHour: Math.max(0, Math.floor(minMinute / 60)),
    endHour: Math.min(24, Math.ceil(maxMinute / 60)),
  };
}

function buildFilters(events) {
  const filters = new Map();

  for (const event of events) {
    if (!filters.has(event.filterKey)) {
      filters.set(event.filterKey, {
        key: event.filterKey,
        label: event.filterLabel,
        kind: event.filterKind,
        subject: event.displaySubject,
        teacher: event.teacher,
        colorKey: event.colorKey,
        sourceColor: event.sourceColor || "",
        type: event.type,
        count: 0,
      });
    }
    filters.get(event.filterKey).count += 1;
  }

  return [...filters.values()].sort((a, b) => {
    const labelCompare = a.label.localeCompare(b.label);
    if (labelCompare) return labelCompare;
    return a.type.label.localeCompare(b.type.label);
  });
}

function toDateKey(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function parseDateKey(value) {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function toLocalIso(date) {
  return `${toDateKey(date)}T${pad(date.getHours())}:${pad(
    date.getMinutes()
  )}:${pad(date.getSeconds())}`;
}

function pad(value) {
  return String(value).padStart(2, "0");
}

function buildIcs(schedule, options = {}) {
  const filterKeys = options.filterKeys || null;
  const events = filterKeys
    ? schedule.events.filter((event) => filterKeys.has(event.filterKey))
    : schedule.events;
  const generated = formatIcsUtc(new Date());
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//fri-urnik//Urnik Viewer//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    "X-WR-CALNAME:FRI urnik",
    `X-WR-TIMEZONE:${TIMEZONE}`,
    ...timezoneComponent(),
  ];

  for (const event of events) {
    lines.push(
      "BEGIN:VEVENT",
      `UID:${escapeIcsValue(`${event.id}@fri-urnik.local`)}`,
      `DTSTAMP:${generated}`,
      `DTSTART;TZID=${TIMEZONE}:${formatIcsLocal(event.start)}`,
      `DTEND;TZID=${TIMEZONE}:${formatIcsLocal(event.end)}`,
      `SUMMARY:${escapeIcsValue(`${event.displaySubject} (${event.type.code})`)}`,
      `LOCATION:${escapeIcsValue(event.location || "")}`,
      `DESCRIPTION:${escapeIcsValue(event.teacher || "")}`,
    );

    const rrule = exportRRule(event);
    if (rrule) {
      lines.push(`RRULE:${rrule}`);
    }

    lines.push("END:VEVENT");
  }

  lines.push("END:VCALENDAR");
  return `${lines.map(foldIcsLine).join("\r\n")}\r\n`;
}

function timezoneComponent() {
  if (TIMEZONE !== "Europe/Ljubljana") {
    return [];
  }

  return [
    "BEGIN:VTIMEZONE",
    "TZID:Europe/Ljubljana",
    "X-LIC-LOCATION:Europe/Ljubljana",
    "BEGIN:DAYLIGHT",
    "TZOFFSETFROM:+0100",
    "TZOFFSETTO:+0200",
    "TZNAME:CEST",
    "DTSTART:19700329T020000",
    "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU",
    "END:DAYLIGHT",
    "BEGIN:STANDARD",
    "TZOFFSETFROM:+0200",
    "TZOFFSETTO:+0100",
    "TZNAME:CET",
    "DTSTART:19701025T030000",
    "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU",
    "END:STANDARD",
    "END:VTIMEZONE",
  ];
}

function exportRRule(event) {
  if (!event.recurrence?.frequency) {
    return "";
  }

  const parts = [`FREQ=${event.recurrence.frequency}`];
  if (event.recurrence.interval && event.recurrence.interval !== 1) {
    parts.push(`INTERVAL=${event.recurrence.interval}`);
  }
  if (event.recurrence.byDay?.length) {
    parts.push(`BYDAY=${event.recurrence.byDay.join(",")}`);
  }
  if (event.lastDate) {
    const until = parseDateKey(event.lastDate);
    until.setHours(23, 59, 59, 0);
    parts.push(`UNTIL=${formatIcsUtc(until)}`);
  }

  return parts.join(";");
}

function formatIcsLocal(value) {
  const date = typeof value === "string" ? parseLocalDateTime(value) : value;
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}T${pad(
    date.getHours()
  )}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}

function formatIcsUtc(date) {
  return `${date.getUTCFullYear()}${pad(date.getUTCMonth() + 1)}${pad(
    date.getUTCDate()
  )}T${pad(date.getUTCHours())}${pad(date.getUTCMinutes())}${pad(
    date.getUTCSeconds()
  )}Z`;
}

function escapeIcsValue(value) {
  return String(value)
    .replace(/\\/g, "\\\\")
    .replace(/\n/g, "\\n")
    .replace(/,/g, "\\,")
    .replace(/;/g, "\\;");
}

function foldIcsLine(line) {
  const max = 74;
  if (Buffer.byteLength(line, "utf8") <= max) {
    return line;
  }

  let output = "";
  let current = "";
  for (const char of line) {
    if (Buffer.byteLength(`${current}${char}`, "utf8") > max) {
      output += `${current}\r\n `;
      current = char;
    } else {
      current += char;
    }
  }
  return output + current;
}

async function serveStatic(req, res, pathname) {
  const cleanPath = pathname === "/" ? "/index.html" : pathname;
  let decoded;

  try {
    decoded = decodeURIComponent(cleanPath);
  } catch {
    sendJson(res, 400, { error: "Invalid path" });
    return;
  }

  const normalized = path.normalize(decoded).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(PUBLIC_DIR, normalized);

  if (!filePath.startsWith(PUBLIC_DIR)) {
    sendJson(res, 403, { error: "Forbidden" });
    return;
  }

  try {
    const data = await fs.readFile(filePath);
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      "content-type": MIME_TYPES[ext] || "application/octet-stream",
      "cache-control": "no-cache",
      ...securityHeaders(),
    });
    res.end(data);
  } catch (error) {
    if (error.code === "ENOENT" || error.code === "EISDIR") {
      sendJson(res, 404, { error: "Not found" });
      return;
    }
    throw error;
  }
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);

    if (url.pathname === "/api/health") {
      sendJson(res, 200, { ok: true });
      return;
    }

    if (url.pathname === "/api/schedule") {
      const sourceUrl = url.searchParams.get("sourceUrl") || DEFAULT_TIMETABLE_URL;
      const forceImport = url.searchParams.get("refresh") === "1";
      const payload = await getSchedule({ sourceUrl, forceImport });
      sendJson(res, 200, payload);
      return;
    }

    if (url.pathname === "/api/config") {
      if (req.method === "GET") {
        const store = await loadConfigStore();
        sendJson(res, 200, {
          sourceUrl: store.sourceUrl || "",
          filters: store.filters || {},
          theme: store.theme || "",
        });
        return;
      }
      if (req.method === "PUT") {
        const body = await readJsonBody(req);
        const store = await loadConfigStore();

        if (Object.prototype.hasOwnProperty.call(body, "sourceUrl")) {
          const raw = typeof body.sourceUrl === "string" ? body.sourceUrl.trim() : "";
          store.sourceUrl = raw ? normalizeTimetableUrl(raw) : "";
        }

        if (Object.prototype.hasOwnProperty.call(body, "filters")) {
          if (body.filters && typeof body.filters === "object") {
            const existing = store.filters || {};
            const incoming = sanitizeFilterStore(body.filters);
            store.filters = { ...existing, ...incoming };
          }
        }

        if (Object.prototype.hasOwnProperty.call(body, "theme")) {
          const t = typeof body.theme === "string" ? body.theme : "";
          store.theme = ["light", "dark"].includes(t) ? t : "";
        }

        await saveConfigStore(store);
        sendJson(res, 200, {
          sourceUrl: store.sourceUrl || "",
          filters: store.filters || {},
          theme: store.theme || "",
        });
        return;
      }
      sendJson(res, 405, { error: "Method not allowed" }, { allow: "GET, PUT" });
      return;
    }

    if (url.pathname === "/api/export.ics") {
      const sourceUrl = url.searchParams.get("sourceUrl") || DEFAULT_TIMETABLE_URL;
      const forceImport = url.searchParams.get("refresh") === "1";
      const filterKeys = url.searchParams.has("filters")
        ? new Set(
            url.searchParams
              .get("filters")
              .split(",")
              .map((key) => key.trim())
              .filter(Boolean)
          )
        : null;
      const schedule = await getSchedule({ sourceUrl, forceImport });
      const ics = buildIcs(schedule, { filterKeys });
      sendText(res, 200, ics, {
        "content-type": "text/calendar; charset=utf-8",
        "content-disposition": 'attachment; filename="urnik.ics"',
      });
      return;
    }

    if (req.method !== "GET" && req.method !== "HEAD") {
      sendJson(res, 405, { error: "Method not allowed" }, { allow: "GET, HEAD" });
      return;
    }

    await serveStatic(req, res, url.pathname);
  } catch (error) {
    const status = error.status || (error.code === "ENOENT" ? 404 : 500);
    sendJson(res, status, {
      error: status === 500 ? "Server error" : error.message,
    });
    if (status === 500) {
      console.error(error);
    }
  }
});

server.listen(PORT, "0.0.0.0", () => {
  console.log(`Urnik viewer listening on http://0.0.0.0:${PORT}`);
});
