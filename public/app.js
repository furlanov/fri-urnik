const COLOR_PALETTE = [
  "hsla(248, 24%, 86%, 0.78)",
  "hsla(6, 62%, 86%, 0.78)",
  "hsla(184, 42%, 84%, 0.78)",
  "hsla(83, 36%, 83%, 0.78)",
  "hsla(42, 70%, 86%, 0.78)",
  "hsla(316, 34%, 88%, 0.78)",
  "hsla(204, 44%, 87%, 0.78)",
  "hsla(132, 28%, 84%, 0.78)",
  "hsla(28, 52%, 85%, 0.78)",
  "hsla(268, 32%, 88%, 0.78)",
];
const DAY_NAMES = ["nedelja", "ponedeljek", "torek", "sreda", "četrtek", "petek", "sobota"];
const SHORT_DAY_NAMES = ["ned", "pon", "tor", "sre", "čet", "pet", "sob"];
const DAY_CODES = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];
const STORAGE_PREFIX = "urnik.filters.v1";
const SOURCE_URL_STORAGE = "urnik.sourceUrl.v1";
const THEME_STORAGE = "urnik.theme.v1";
const THEMES = ["light", "dark"];
const MIN_EVENT_HEIGHT = 38;
const HOUR_HEIGHT = 60;

const state = {
  schedule: null,
  weekStart: null,
  filters: {},
  search: "",
  sourceUrl: safeGetLocalStorage(SOURCE_URL_STORAGE),
  theme: readStoredTheme(),
  serverConfig: { sourceUrl: "", filters: {}, theme: "" },
  serverConfigLoaded: false,
};

applyTheme(state.theme);

const elements = {
  sourceUrlInput: document.querySelector("#sourceUrlInput"),
  saveSourceUrl: document.querySelector("#saveSourceUrl"),
  refreshImport: document.querySelector("#refreshImport"),
  exportIcs: document.querySelector("#exportIcs"),
  importStatus: document.querySelector("#importStatus"),
  filterList: document.querySelector("#filterList"),
  filterSearch: document.querySelector("#filterSearch"),
  previousWeek: document.querySelector("#previousWeek"),
  currentWeek: document.querySelector("#currentWeek"),
  nextWeek: document.querySelector("#nextWeek"),
  showAll: document.querySelector("#showAll"),
  hideAll: document.querySelector("#hideAll"),
  lecturesOnly: document.querySelector("#lecturesOnly"),
  labsOnly: document.querySelector("#labsOnly"),
  openSettings: document.querySelector("#openSettings"),
  closeSettings: document.querySelector("#closeSettings"),
  settingsBackdrop: document.querySelector("#settingsBackdrop"),
  settingsPanel: document.querySelector("#settingsPanel"),
  weekTitle: document.querySelector("#weekTitle"),
  brandWeek: document.querySelector("#brandWeek"),
  themeSwitch: document.querySelector("#themeSwitch"),
  programContext: document.querySelector("#programContext"),
  mobileSchedule: document.querySelector("#mobileSchedule"),
  timetable: document.querySelector("#timetable"),
  emptyState: document.querySelector("#emptyState"),
};

let nowLineTimer = null;
let configWriteTimer = null;
let pendingConfigPatch = null;
let shouldScrollToToday = false;

init();

async function init() {
  elements.sourceUrlInput.value = state.sourceUrl;
  elements.previousWeek.addEventListener("click", () => shiftWeek(-1));
  elements.nextWeek.addEventListener("click", () => shiftWeek(1));
  elements.currentWeek.addEventListener("click", () => {
    if (!state.schedule) return;
    state.weekStart = chooseDefaultWeek(state.schedule);
    shouldScrollToToday = true;
    render();
  });
  elements.saveSourceUrl.addEventListener("click", () => {
    saveSourceUrl();
    loadSchedule({ sourceUrl: state.sourceUrl });
  });
  elements.refreshImport.addEventListener("click", () => {
    saveSourceUrl();
    loadSchedule({ sourceUrl: state.sourceUrl, refresh: true });
  });
  elements.exportIcs.addEventListener("click", exportVisibleCalendar);
  elements.sourceUrlInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      saveSourceUrl();
      loadSchedule({ sourceUrl: state.sourceUrl });
    }
  });
  elements.filterSearch.addEventListener("input", () => {
    state.search = elements.filterSearch.value.trim().toLowerCase();
    renderFilters();
  });
  elements.showAll.addEventListener("click", () => setFilters(() => true));
  elements.hideAll.addEventListener("click", () => setFilters(() => false));
  elements.lecturesOnly.addEventListener("click", () =>
    setFilters((filter) => filter.type.role === "lecture")
  );
  elements.labsOnly.addEventListener("click", () =>
    setFilters((filter) => filter.type.role === "lab")
  );
  if (elements.themeSwitch) {
    elements.themeSwitch.addEventListener("click", (event) => {
      const target = event.target.closest("button[data-theme-value]");
      if (!target) return;
      const value = target.dataset.themeValue;
      if (!THEMES.includes(value)) return;
      setTheme(value, { persist: true });
    });
    syncThemeSwitch();
  }

  elements.openSettings.addEventListener("click", openSettings);
  elements.closeSettings.addEventListener("click", closeSettings);
  elements.settingsBackdrop.addEventListener("click", closeSettings);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeSettings();
    if (!state.schedule) return;
    if ((event.target instanceof HTMLInputElement) || (event.target instanceof HTMLTextAreaElement)) return;
    if (event.key === "ArrowLeft") shiftWeek(-1);
    if (event.key === "ArrowRight") shiftWeek(1);
  });

  attachSwipeNavigation(elements.mobileSchedule);
  trackHeaderHeight();
  window.addEventListener("pagehide", flushConfigSync);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") flushConfigSync();
  });

  await loadServerConfig();
  elements.sourceUrlInput.value = state.sourceUrl;
  loadSchedule({ sourceUrl: state.sourceUrl });
}

function attachSwipeNavigation(target) {
  if (!target) return;
  let startX = 0;
  let startY = 0;
  let startTime = 0;
  let tracking = false;

  target.addEventListener(
    "touchstart",
    (event) => {
      if (event.touches.length !== 1) return;
      const touch = event.touches[0];
      startX = touch.clientX;
      startY = touch.clientY;
      startTime = Date.now();
      tracking = true;
    },
    { passive: true }
  );

  target.addEventListener(
    "touchend",
    (event) => {
      if (!tracking) return;
      tracking = false;
      const touch = event.changedTouches[0];
      if (!touch) return;
      const dx = touch.clientX - startX;
      const dy = touch.clientY - startY;
      const dt = Date.now() - startTime;
      if (Math.abs(dx) < 60 || Math.abs(dx) < Math.abs(dy) * 1.5 || dt > 600) return;
      shiftWeek(dx > 0 ? -1 : 1);
    },
    { passive: true }
  );

  target.addEventListener("touchcancel", () => {
    tracking = false;
  });
}

function trackHeaderHeight() {
  const header = document.querySelector(".header");
  if (!header) return;
  const apply = () => {
    document.documentElement.style.setProperty(
      "--header-height",
      `${Math.round(header.getBoundingClientRect().height)}px`
    );
  };
  apply();
  if ("ResizeObserver" in window) {
    new ResizeObserver(apply).observe(header);
  } else {
    window.addEventListener("resize", apply);
  }
}

function flushConfigSync() {
  if (!pendingConfigPatch) return;
  const body = pendingConfigPatch;
  pendingConfigPatch = null;
  clearTimeout(configWriteTimer);
  try {
    fetch("/api/config", {
      method: "PUT",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(body),
      keepalive: true,
    });
  } catch {
    // Tab is unloading; best effort only.
  }
}

async function loadServerConfig() {
  try {
    const response = await fetch("/api/config", { headers: { accept: "application/json" } });
    if (!response.ok) return;
    const config = await response.json();
    state.serverConfig = {
      sourceUrl: config?.sourceUrl || "",
      filters: config?.filters && typeof config.filters === "object" ? config.filters : {},
      theme: THEMES.includes(config?.theme) ? config.theme : "",
    };
    state.serverConfigLoaded = true;
    if (state.serverConfig.sourceUrl) {
      state.sourceUrl = state.serverConfig.sourceUrl;
      safeSetLocalStorage(SOURCE_URL_STORAGE, state.sourceUrl);
    }
    if (state.serverConfig.theme && state.serverConfig.theme !== state.theme) {
      setTheme(state.serverConfig.theme, { persist: false });
    }
  } catch {
    // Server unreachable — continue with localStorage fallback.
  }
}

function queueConfigSave(patch) {
  if (!pendingConfigPatch) pendingConfigPatch = {};
  if (Object.prototype.hasOwnProperty.call(patch, "sourceUrl")) {
    pendingConfigPatch.sourceUrl = patch.sourceUrl;
  }
  if (Object.prototype.hasOwnProperty.call(patch, "theme")) {
    pendingConfigPatch.theme = patch.theme;
  }
  if (patch.filters) {
    pendingConfigPatch.filters = {
      ...(pendingConfigPatch.filters || {}),
      ...patch.filters,
    };
  }
  clearTimeout(configWriteTimer);
  configWriteTimer = setTimeout(flushConfig, 400);
}

async function flushConfig() {
  if (!pendingConfigPatch) return;
  const body = pendingConfigPatch;
  pendingConfigPatch = null;
  try {
    const response = await fetch("/api/config", {
      method: "PUT",
      headers: { "content-type": "application/json", accept: "application/json" },
      body: JSON.stringify(body),
    });
    if (!response.ok) return;
    const updated = await response.json();
    state.serverConfig = {
      sourceUrl: updated?.sourceUrl || "",
      filters: updated?.filters && typeof updated.filters === "object" ? updated.filters : {},
      theme: THEMES.includes(updated?.theme) ? updated.theme : "",
    };
  } catch {
    // Queue lost; local state remains consistent via localStorage fallback.
  }
}

function saveSourceUrl() {
  state.sourceUrl = elements.sourceUrlInput.value.trim();
  if (state.sourceUrl) {
    safeSetLocalStorage(SOURCE_URL_STORAGE, state.sourceUrl);
  } else {
    safeRemoveLocalStorage(SOURCE_URL_STORAGE);
  }
  queueConfigSave({ sourceUrl: state.sourceUrl });
}

function openSettings() {
  elements.settingsPanel.hidden = false;
  elements.settingsBackdrop.hidden = false;
  document.body.classList.add("settings-open");
  elements.openSettings.setAttribute("aria-expanded", "true");
  requestAnimationFrame(() => {
    elements.settingsPanel.classList.add("open");
    elements.settingsPanel.setAttribute("aria-hidden", "false");
    elements.closeSettings.focus();
  });
}

function closeSettings() {
  elements.settingsPanel.classList.remove("open");
  elements.settingsPanel.setAttribute("aria-hidden", "true");
  elements.openSettings.setAttribute("aria-expanded", "false");
  document.body.classList.remove("settings-open");
  elements.settingsBackdrop.hidden = true;
  elements.settingsPanel.hidden = true;
}

async function loadSchedule(options = {}) {
  try {
    setLoading(true);
    const params = new URLSearchParams();
    const previousSourceUrl = state.sourceUrl;
    const sourceUrl = options.sourceUrl ?? state.sourceUrl;
    if (sourceUrl) params.set("sourceUrl", sourceUrl);
    if (options.refresh) params.set("refresh", "1");
    const suffix = params.toString() ? `?${params}` : "";
    const response = await fetch(`/api/schedule${suffix}`, {
      headers: { accept: "application/json" },
    });
    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      throw new Error(body.error || `Request failed with ${response.status}`);
    }

    state.schedule = await response.json();
    state.sourceUrl = state.schedule.source?.url || state.schedule.enrichment?.url || sourceUrl || "";
    if (state.sourceUrl) {
      safeSetLocalStorage(SOURCE_URL_STORAGE, state.sourceUrl);
    }
    elements.sourceUrlInput.value = state.sourceUrl;
    state.weekStart = chooseDefaultWeek(state.schedule);
    shouldScrollToToday = true;
    hydrateFilters([sourceUrl, previousSourceUrl]);
    renderProgramContext();
    render();
    updateImportStatus();
  } catch (error) {
    elements.weekTitle.textContent = "";
    elements.importStatus.textContent = error.message;
    elements.emptyState.hidden = false;
    elements.emptyState.textContent = error.message;
  } finally {
    setLoading(false);
  }
}

function renderProgramContext() {
  const context = state.schedule?.context;
  if (!elements.programContext) return;
  const text = context?.subtitle || context?.title || "";
  elements.programContext.textContent = text;
  elements.programContext.title = text;
}

function updateImportStatus() {
  const enrichment = state.schedule?.enrichment;
  if (!enrichment?.enabled) {
    elements.importStatus.textContent = "Vnesi povezavo do FRI urnika.";
    return;
  }

  if (enrichment.error) {
    elements.importStatus.textContent = `Predpomnjeni podatki: ${enrichment.error}`;
    return;
  }

  const checked = enrichment.fetchedAt ? formatDateTime(enrichment.fetchedAt) : "ni še preverjeno";
  elements.importStatus.textContent =
    enrichment.source === "cache"
      ? `Predpomnjenih ${enrichment.applied} terminov. Zadnje uspešno preverjanje: ${checked}.`
      : `Uvoženih ${enrichment.applied} terminov. Zadnje dnevno preverjanje: ${checked}.`;
}

function setLoading(isLoading) {
  document.body.classList.toggle("is-loading", isLoading);
}

function exportVisibleCalendar() {
  saveSourceUrl();
  const params = new URLSearchParams();
  if (state.sourceUrl) {
    params.set("sourceUrl", state.sourceUrl);
  }
  if (state.schedule) {
    const activeFilters = Object.entries(state.filters)
      .filter(([, enabled]) => enabled)
      .map(([key]) => key);
    params.set("filters", activeFilters.join(","));
  }

  window.location.href = `/api/export.ics?${params}`;
}

function chooseDefaultWeek(schedule) {
  const today = startOfDay(new Date());
  const rangeStart = schedule?.range?.start ? parseDateKey(schedule.range.start) : today;
  const rangeEnd = schedule?.range?.end ? parseDateKey(schedule.range.end) : today;

  if (today >= rangeStart && today <= rangeEnd) {
    const day = today.getDay();
    const reference = day === 0 || day === 6 ? addDays(today, 8 - (day || 7)) : today;
    return getWeekStart(reference);
  }
  return getWeekStart(rangeStart);
}

function hydrateFilters(sourceCandidates = []) {
  const saved = readSavedFilters(sourceCandidates);
  state.filters = {};
  for (const filter of state.schedule.filters) {
    state.filters[filter.key] = Object.prototype.hasOwnProperty.call(saved, filter.key)
      ? Boolean(saved[filter.key])
      : true;
  }
  persistFilters({ push: false });
}

function readSavedFilters(sourceCandidates = []) {
  const canonicalCandidates = [
    state.schedule?.source?.url,
    state.schedule?.enrichment?.url,
    state.sourceUrl,
    ...sourceCandidates,
  ].filter(Boolean);

  const serverFilters = state.serverConfig?.filters || {};
  for (const candidate of canonicalCandidates) {
    const stored = serverFilters[candidate];
    if (stored && typeof stored === "object" && Object.keys(stored).length > 0) {
      return stored;
    }
  }

  for (const key of filterStorageKeys(sourceCandidates)) {
    try {
      const saved = JSON.parse(safeGetLocalStorage(key) || "{}");
      if (saved && typeof saved === "object" && Object.keys(saved).length > 0) {
        return saved;
      }
    } catch {
      // Ignore malformed or inaccessible localStorage entries.
    }
  }
  return {};
}

function persistFilters(options = {}) {
  try {
    safeSetLocalStorage(storageKey(), JSON.stringify(state.filters));
  } catch {
    // Safari private mode and locked-down browsers can reject localStorage writes.
  }

  if (options.push === false) return;
  const canonical = state.schedule?.source?.url || state.sourceUrl;
  if (!canonical) return;
  queueConfigSave({ filters: { [canonical]: state.filters } });
}

function storageKey() {
  return `${STORAGE_PREFIX}:${sourceStorageId(
    state.schedule?.source?.url || state.sourceUrl || "default"
  )}`;
}

function filterStorageKeys(sourceCandidates = []) {
  const keys = new Set([storageKey()]);
  const storedSourceUrl = safeGetLocalStorage(SOURCE_URL_STORAGE);
  const sources = [
    ...sourceCandidates,
    state.sourceUrl,
    state.schedule?.source?.url,
    state.schedule?.enrichment?.url,
    storedSourceUrl,
  ].filter(Boolean);

  for (const source of sources) {
    keys.add(`${STORAGE_PREFIX}:${sourceStorageId(source)}`);
    keys.add(legacyStorageKey(source));
  }

  return [...keys];
}

function sourceStorageId(source) {
  const value = String(source || "default").trim();
  if (!value || value === "default") return "default";

  try {
    const url = new URL(value, window.location.origin);
    const group = url.searchParams.get("group");
    const timetable = url.pathname.match(/\/timetable\/([^/]+)\//)?.[1];
    if (group && timetable) {
      return `fri:${timetable}:group:${group}`;
    }

    const params = Array.from(url.searchParams.entries()).sort(([keyA, valueA], [keyB, valueB]) =>
      keyA === keyB ? valueA.localeCompare(valueB) : keyA.localeCompare(keyB)
    );
    url.hash = "";
    url.search = "";
    for (const [key, paramValue] of params) {
      url.searchParams.append(key, paramValue);
    }
    return encodeURIComponent(url.toString());
  } catch {
    return encodeURIComponent(value);
  }
}

function legacyStorageKey(source) {
  return `${STORAGE_PREFIX}:${encodeURIComponent(String(source || "default").trim() || "default")}`;
}

function readStoredTheme() {
  const stored = safeGetLocalStorage(THEME_STORAGE);
  if (THEMES.includes(stored)) return stored;
  const prefersDark =
    window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  return prefersDark ? "dark" : "light";
}

function setTheme(value, options = {}) {
  const next = THEMES.includes(value) ? value : "light";
  state.theme = next;
  applyTheme(next);
  syncThemeSwitch();
  if (options.persist === false) return;
  safeSetLocalStorage(THEME_STORAGE, next);
  queueConfigSave({ theme: next });
}

function applyTheme(value) {
  const root = document.documentElement;
  root.dataset.theme = value;
  const isDark = value === "dark";
  root.classList.toggle("is-dark", isDark);
  const themeColor = document.querySelector('meta[name="theme-color"]');
  if (themeColor) themeColor.setAttribute("content", isDark ? "#0d1512" : "#0b8f78");
}

function syncThemeSwitch() {
  if (!elements.themeSwitch) return;
  for (const button of elements.themeSwitch.querySelectorAll("button[data-theme-value]")) {
    const active = button.dataset.themeValue === state.theme;
    button.setAttribute("aria-checked", active ? "true" : "false");
  }
}

function safeGetLocalStorage(key) {
  try { return localStorage.getItem(key) || ""; } catch { return ""; }
}

function safeSetLocalStorage(key, value) {
  try { localStorage.setItem(key, value); } catch {}
}

function safeRemoveLocalStorage(key) {
  try { localStorage.removeItem(key); } catch {}
}

function setFilters(predicate) {
  for (const filter of state.schedule.filters) {
    state.filters[filter.key] = Boolean(predicate(filter));
  }
  persistFilters();
  render();
}

function shiftWeek(delta) {
  state.weekStart = addDays(state.weekStart, delta * 7);
  render();
}

function render() {
  if (!state.schedule) return;

  const allOccurrences = expandWeek(state.schedule.events, state.weekStart);
  const visibleOccurrences = allOccurrences.filter((event) => state.filters[event.filterKey]);
  const days = displayDaysFor(allOccurrences);
  const visibleDays = days.map((dayIndex) => {
    const date = addDays(state.weekStart, dayIndex === 0 ? 6 : dayIndex - 1);
    return {
      dayIndex,
      date,
      workFreeDay: holidayForDate(date),
      events: visibleOccurrences.filter((event) => event.weekday === dayIndex),
    };
  });

  renderFilters();
  renderHeader();
  renderTimetable(visibleDays, visibleOccurrences);
  renderMobileSchedule(visibleDays);

  const hasVisibleEvents = visibleOccurrences.length > 0;
  elements.emptyState.hidden = hasVisibleEvents;

  scheduleNowLineRefresh();
}

function renderHeader() {
  const weekEnd = addDays(state.weekStart, 6);
  elements.weekTitle.textContent = `${formatDate(state.weekStart)} – ${formatDate(weekEnd)}`;
  if (elements.brandWeek) {
    elements.brandWeek.textContent = formatCompactWeek(state.weekStart, weekEnd);
  }
}

function formatCompactWeek(start, end) {
  const startDay = start.getDate();
  const endDay = end.getDate();
  const monthFmt = new Intl.DateTimeFormat("sl-SI", { month: "short" });
  if (start.getMonth() === end.getMonth()) {
    return `${startDay}.–${endDay}. ${monthFmt.format(end)}`;
  }
  return `${startDay}. ${monthFmt.format(start)} – ${endDay}. ${monthFmt.format(end)}`;
}

function renderFilters() {
  const groups = groupFilters(state.schedule.filters);

  const fragment = document.createDocumentFragment();
  for (const group of groups) {
    if (state.search && !group.label.toLowerCase().includes(state.search)) {
      continue;
    }
    fragment.appendChild(createFilterGroup(group));
  }

  elements.filterList.replaceChildren(fragment);
}

function createFilterGroup(group) {
  const wrapper = document.createElement("section");
  wrapper.className = "filter-group";
  wrapper.style.setProperty("--subject-color", colorFor(group.colorKey, group.sourceColor));

  const title = document.createElement("div");
  title.className = "filter-title";
  const dot = document.createElement("span");
  dot.className = "color-dot";
  const text = document.createElement("span");
  text.className = "label";
  text.textContent = group.label;
  title.append(dot, text);
  if (group.kind === "instructor") {
    const kind = document.createElement("small");
    kind.className = "filter-kind";
    kind.textContent = "Izvajalec";
    title.appendChild(kind);
  }

  const row = document.createElement("div");
  row.className = "toggle-row";

  for (const filter of group.filters) {
    const label = document.createElement("label");
    label.className = "toggle";

    const input = document.createElement("input");
    input.type = "checkbox";
    input.checked = Boolean(state.filters[filter.key]);
    input.addEventListener("change", () => {
      state.filters[filter.key] = input.checked;
      persistFilters();
      render();
    });

    const labelText = document.createElement("span");
    labelText.textContent = filter.type.label;
    const count = document.createElement("small");
    count.textContent = String(filter.count);
    label.append(input, labelText, count);
    row.appendChild(label);
  }

  wrapper.append(title, row);
  return wrapper;
}

function groupFilters(filters) {
  const byLabel = new Map();
  for (const filter of filters) {
    if (!byLabel.has(filter.label)) {
      byLabel.set(filter.label, {
        label: filter.label,
        kind: filter.kind,
        colorKey: filter.colorKey,
        sourceColor: filter.sourceColor || "",
        filters: [],
      });
    }
    byLabel.get(filter.label).filters.push(filter);
  }

  return [...byLabel.values()].sort((a, b) => a.label.localeCompare(b.label));
}

function renderTimetable(days, visibleOccurrences) {
  const startHour = Math.max(0, state.schedule.range.startHour ?? 7);
  const endHour = Math.max(startHour + 1, state.schedule.range.endHour ?? 18);
  const hourCount = endHour - startHour;
  const gridHeight = hourCount * HOUR_HEIGHT;
  const laidOutDays = days.map((day) => ({
    ...day,
    events: layoutOverlaps(day.events),
  }));

  elements.timetable.style.setProperty("--day-count", String(days.length));
  elements.timetable.style.setProperty("--hour-count", String(hourCount));
  elements.timetable.style.setProperty("--grid-height", `${gridHeight}px`);
  elements.timetable.dataset.startHour = String(startHour);
  elements.timetable.replaceChildren();

  for (let hour = startHour; hour <= endHour; hour += 1) {
    const label = document.createElement("div");
    label.className = "grid-hour";
    label.style.gridRow = `${hour - startHour + 2}`;
    label.textContent = `${String(hour).padStart(2, "0")}:00`;
    elements.timetable.appendChild(label);
  }

  for (let hour = 0; hour < hourCount; hour += 1) {
    const row = document.createElement("div");
    row.className = "grid-complete-row";
    row.style.gridRow = `${hour + 2} / span 1`;
    elements.timetable.appendChild(row);
  }

  const today = new Date();
  for (const [index, day] of laidOutDays.entries()) {
    elements.timetable.appendChild(dayHeader(day, index + 2, today));
  }

  for (const [index, day] of laidOutDays.entries()) {
    const column = document.createElement("div");
    const isToday = isSameDate(day.date, today);
    column.className = `grid-day-column${isToday ? " today" : ""}${day.workFreeDay ? " holiday" : ""}`;
    column.style.gridColumn = String(index + 2);
    if (isToday) column.dataset.today = "1";
    if (day.workFreeDay) {
      column.title = `Dela prost dan: ${day.workFreeDay.name}`;
    }

    for (const event of day.events) {
      const top = ((event.startMinute - startHour * 60) / 60) * HOUR_HEIGHT;
      const height = Math.max(
        MIN_EVENT_HEIGHT,
        (event.durationMinutes / 60) * HOUR_HEIGHT - 6
      );
      const card = createEventCard(event);
      card.style.setProperty("--event-top", `${top}px`);
      card.style.setProperty("--event-height", `${height}px`);
      card.style.setProperty("--event-lane", String(event.lane));
      card.style.setProperty("--event-lanes", String(event.lanes));
      column.appendChild(card);
    }

    elements.timetable.appendChild(column);
  }

  elements.timetable.hidden = visibleOccurrences.length === 0;

  updateNowLine();
}

function scheduleNowLineRefresh() {
  if (nowLineTimer) clearInterval(nowLineTimer);
  nowLineTimer = setInterval(updateNowLine, 60000);
}

function updateNowLine() {
  const existing = elements.timetable.querySelector(".now-line");
  if (existing) existing.remove();
  const column = elements.timetable.querySelector(".grid-day-column[data-today]");
  if (!column) return;

  const startHour = Number(elements.timetable.dataset.startHour || "7");
  const now = new Date();
  const minutesFromStart = now.getHours() * 60 + now.getMinutes() - startHour * 60;
  if (minutesFromStart < 0) return;
  const top = (minutesFromStart / 60) * HOUR_HEIGHT;
  if (top > parseFloat(elements.timetable.style.getPropertyValue("--grid-height"))) return;

  const line = document.createElement("div");
  line.className = "now-line";
  line.style.top = `${top}px`;
  column.appendChild(line);
}

function renderMobileSchedule(days) {
  const fragment = document.createDocumentFragment();
  const today = new Date();
  let todaySection = null;

  for (const day of days) {
    const section = document.createElement("section");
    section.className = "mobile-day";
    const isToday = isSameDate(day.date, today);
    if (isToday) {
      section.classList.add("is-today");
      todaySection = section;
    }

    const header = document.createElement("div");
    header.className = "mobile-day-header";
    if (isToday) header.classList.add("is-today");
    if (day.workFreeDay) {
      header.classList.add("is-holiday");
      header.title = `Dela prost dan: ${day.workFreeDay.name}`;
    }
    const name = document.createElement("span");
    name.className = "day-name";
    name.textContent = DAY_NAMES[day.dayIndex];
    const date = document.createElement("span");
    date.className = "day-date";
    date.textContent = formatDate(day.date);
    header.append(name, date);
    if (day.workFreeDay) {
      header.appendChild(createHolidayChip(day.workFreeDay));
    }
    section.appendChild(header);

    const list = document.createElement("div");
    list.className = "mobile-day-list";

    if (day.events.length) {
      for (const event of day.events) {
        list.appendChild(createMobileEntry(event));
      }
    } else {
      const empty = document.createElement("p");
      empty.className = "mobile-day-empty";
      empty.textContent = "Ni terminov";
      list.appendChild(empty);
    }

    section.appendChild(list);
    fragment.appendChild(section);
  }

  elements.mobileSchedule.replaceChildren(fragment);

  if (shouldScrollToToday && todaySection && window.matchMedia("(max-width: 900px)").matches) {
    requestAnimationFrame(() => {
      const headerOffset = parseFloat(
        getComputedStyle(document.documentElement).getPropertyValue("--header-height")
      ) || 0;
      const top =
        todaySection.getBoundingClientRect().top + window.scrollY - headerOffset - 4;
      window.scrollTo({ top: Math.max(0, top), behavior: "auto" });
    });
  }
  shouldScrollToToday = false;
}

function dayHeader(day, gridColumn, today) {
  const node = document.createElement("div");
  const isToday = isSameDate(day.date, today);
  node.className = `grid-day${isToday ? " is-today" : ""}${day.workFreeDay ? " is-holiday" : ""}`;
  node.style.gridColumn = String(gridColumn);
  if (day.workFreeDay) {
    node.title = `Dela prost dan: ${day.workFreeDay.name}`;
  }

  const wrap = document.createElement("div");
  wrap.className = "grid-day-label";
  const name = document.createElement("span");
  name.textContent = SHORT_DAY_NAMES[day.dayIndex];
  const date = document.createElement("span");
  date.className = "day-date";
  date.textContent = formatDate(day.date);
  wrap.append(name, date);
  if (day.workFreeDay) {
    wrap.appendChild(createHolidayChip(day.workFreeDay));
  }
  node.appendChild(wrap);
  return node;
}

function createHolidayChip(workFreeDay) {
  const chip = document.createElement("span");
  chip.className = "holiday-chip";
  chip.textContent = workFreeDay.name;
  chip.title = `Dela prost dan: ${workFreeDay.name}`;
  return chip;
}

function createEventCard(event) {
  const card = document.createElement("article");
  card.className = "grid-entry";
  card.tabIndex = 0;
  card.dataset.role = event.type.role;
  if (event.lanes >= 3) {
    card.classList.add("is-cramped");
  } else if (event.lanes === 2) {
    card.classList.add("is-tight");
  }
  const color = colorFor(event.colorKey, event.sourceColor);
  card.style.setProperty("--subject-color", color);
  card.style.setProperty("--entry-color", color);
  const timeRange = `${formatTime(event.start)}–${formatTime(event.end)}`;
  const visibleTimeRange =
    event.lanes >= 3
      ? `${formatCompactTime(event.start)}–${formatCompactTime(event.end)}`
      : timeRange;

  const description = document.createElement("div");
  description.className = "description";

  const subjectRow = document.createElement("div");
  subjectRow.className = "row subject-row";
  const subject = document.createElement("span");
  subject.className = "link-subject";
  subject.textContent = event.displaySubject;
  subjectRow.appendChild(subject);

  const metaRow = document.createElement("div");
  metaRow.className = "row meta-row";
  const time = document.createElement("span");
  time.className = "time-range";
  time.textContent = visibleTimeRange;
  metaRow.appendChild(time);
  if (event.location) {
    const room = document.createElement("span");
    room.className = "link-classroom";
    room.textContent = event.location;
    metaRow.appendChild(room);
  }

  description.appendChild(subjectRow);
  description.appendChild(metaRow);

  if (event.teacher) {
    const teacherRow = document.createElement("div");
    teacherRow.className = "row teacher-row";
    const teacher = document.createElement("span");
    teacher.className = "link-teacher";
    teacher.textContent = event.teacher;
    teacherRow.appendChild(teacher);
    description.appendChild(teacherRow);
  }

  card.appendChild(description);
  return card;
}

function createMobileEntry(event) {
  const wrap = document.createElement("article");
  wrap.className = "mobile-entry";
  wrap.dataset.role = event.type.role;
  wrap.tabIndex = 0;

  const timeCol = document.createElement("div");
  timeCol.className = "time-col";
  const start = document.createElement("span");
  start.className = "start";
  start.textContent = formatTime(event.start);
  const end = document.createElement("span");
  end.className = "end";
  end.textContent = formatTime(event.end);
  timeCol.append(start, end);

  const card = document.createElement("div");
  card.className = "card";
  card.style.setProperty("--entry-color", colorFor(event.colorKey, event.sourceColor));

  const titleRow = document.createElement("div");
  titleRow.className = "title-row";
  const subject = document.createElement("span");
  subject.className = "subject";
  subject.textContent = event.displaySubject;
  titleRow.appendChild(subject);

  const meta = document.createElement("div");
  meta.className = "meta";
  if (event.location) {
    const room = document.createElement("span");
    room.className = "room";
    room.textContent = event.location;
    meta.appendChild(room);
  }
  if (event.teacher) {
    const teacher = document.createElement("span");
    teacher.className = "teacher";
    teacher.textContent = event.teacher;
    meta.appendChild(teacher);
  }

  card.append(titleRow, meta);
  wrap.append(timeCol, card);
  return wrap;
}

function expandWeek(events, weekStart) {
  const occurrences = [];
  for (const event of events) {
    for (let offset = 0; offset < 7; offset += 1) {
      const date = addDays(weekStart, offset);
      if (!occursOn(event, date)) continue;
      occurrences.push(toOccurrence(event, date));
    }
  }
  return occurrences.sort((a, b) => a.weekday - b.weekday || a.startMinute - b.startMinute);
}

function occursOn(event, date) {
  const dateKey = toDateKey(date);
  if (dateKey < event.firstDate || dateKey > event.lastDate) {
    return false;
  }

  if (!event.recurrence) {
    return dateKey === event.firstDate;
  }

  if (event.recurrence.frequency !== "WEEKLY") {
    return dateKey === event.firstDate;
  }

  if (!event.recurrence.byDay.includes(DAY_CODES[date.getDay()])) {
    return false;
  }

  const firstWeek = getWeekStart(parseDateKey(event.firstDate));
  const currentWeek = getWeekStart(date);
  const diffWeeks = Math.round((currentWeek - firstWeek) / (7 * 24 * 60 * 60 * 1000));
  return diffWeeks >= 0 && diffWeeks % event.recurrence.interval === 0;
}

function toOccurrence(event, date) {
  const start = parseLocalDateTime(event.start);
  const end = parseLocalDateTime(event.end);
  const occurrenceStart = new Date(date);
  occurrenceStart.setHours(start.getHours(), start.getMinutes(), start.getSeconds(), 0);
  const occurrenceEnd = new Date(date);
  occurrenceEnd.setHours(end.getHours(), end.getMinutes(), end.getSeconds(), 0);

  return {
    ...event,
    occurrenceDate: toDateKey(date),
    start: toLocalIso(occurrenceStart),
    end: toLocalIso(occurrenceEnd),
    weekday: date.getDay(),
  };
}

function layoutOverlaps(events) {
  const sorted = [...events].sort(
    (a, b) => a.startMinute - b.startMinute || b.endMinute - a.endMinute
  );
  const groups = [];
  let group = [];
  let groupEnd = -1;

  for (const event of sorted) {
    if (!group.length || event.startMinute < groupEnd) {
      group.push(event);
      groupEnd = Math.max(groupEnd, event.endMinute);
    } else {
      groups.push(group);
      group = [event];
      groupEnd = event.endMinute;
    }
  }
  if (group.length) groups.push(group);

  return groups.flatMap(assignLanes);
}

function assignLanes(group) {
  const active = [];
  let laneCount = 0;
  const assigned = group.map((event) => {
    for (let index = active.length - 1; index >= 0; index -= 1) {
      if (active[index].endMinute <= event.startMinute) {
        active.splice(index, 1);
      }
    }

    let lane = 0;
    while (active.some((item) => item.lane === lane)) {
      lane += 1;
    }

    const copy = { ...event, lane };
    active.push(copy);
    laneCount = Math.max(laneCount, lane + 1);
    return copy;
  });

  return assigned.map((event) => ({ ...event, lanes: laneCount }));
}

function displayDaysFor(occurrences) {
  const hasWeekend = occurrences.some((event) => event.weekday === 0 || event.weekday === 6);
  return hasWeekend ? [1, 2, 3, 4, 5, 6, 0] : [1, 2, 3, 4, 5];
}

function holidayForDate(date) {
  return state.schedule?.holidays?.dates?.[toDateKey(date)] || null;
}

function colorFor(key, sourceColor) {
  if (sourceColor) return sourceColor;
  const value = [...(key || "")].reduce((sum, char) => sum + char.charCodeAt(0), 0);
  return COLOR_PALETTE[value % COLOR_PALETTE.length];
}

function getWeekStart(date) {
  const copy = startOfDay(date);
  const day = copy.getDay() || 7;
  copy.setDate(copy.getDate() - day + 1);
  return copy;
}

function addDays(date, days) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

function startOfDay(date) {
  const copy = new Date(date);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

function parseDateKey(value) {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function parseLocalDateTime(value) {
  const [datePart, timePart] = value.split("T");
  const [year, month, day] = datePart.split("-").map(Number);
  const [hour, minute, second = 0] = timePart.split(":").map(Number);
  return new Date(year, month - 1, day, hour, minute, second);
}

function toDateKey(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function toLocalIso(date) {
  return `${toDateKey(date)}T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(
    date.getSeconds()
  )}`;
}

function pad(value) {
  return String(value).padStart(2, "0");
}

function formatDate(date) {
  return new Intl.DateTimeFormat("sl-SI", {
    day: "numeric",
    month: "short",
  }).format(date);
}

function formatDateTime(value) {
  return new Intl.DateTimeFormat("sl-SI", {
    day: "2-digit",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(value));
}

function formatTime(value) {
  const date = parseLocalDateTime(value);
  return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function formatCompactTime(value) {
  const date = parseLocalDateTime(value);
  if (date.getMinutes() === 0) {
    return String(date.getHours());
  }
  return `${date.getHours()}:${pad(date.getMinutes())}`;
}

function isSameDate(a, b) {
  return toDateKey(a) === toDateKey(b);
}
