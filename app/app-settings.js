// app-settings.js — wersja w menu ⋯, „Co nowego” po aktualizacji i kopia ustawień (.json).
// (audyt QoL 2026-11-06: S12, S13; kontrakt UX z Documents Workbench pkt 5)

// ── S12: wersja i „Co nowego” ────────────────────────────────────────────────
// Krótko, językiem korzyści — pokazywane raz po aktualizacji (dymek z „Co nowego”)
// i zawsze z menu ⋯ (klik w wersję → dymek „Skopiowano” z tą samą akcją).
const SWB_WHATS_NEW = {
  pl: [
    "Motyw: Systemowy / Jasny / Ciemny — „Systemowy” zmienia się razem z ustawieniem telefonu lub komputera.",
    "Pytania jak w Excelu: „Zapisz / Nie zapisuj / Anuluj” zamiast systemowego OK/Anuluj; nazwy widoków i presetów wpisujesz we własnym oknie.",
    "Po wklejeniu, wypełnieniu, „Zamień wszystkie”, wyczyszczeniu czy przeniesieniu komórek — „Cofnij” prosto w dymku.",
    "Ctrl/⌘ + / otwiera wyszukiwalną listę skrótów; motyw szybko przełączysz Ctrl/⌘ + Alt + /.",
    "Po ponownym otwarciu tego samego pliku aplikacja proponuje powrót do arkusza i komórki, na której skończyłeś.",
    "Kopia ustawień (.json) w menu ⋯ — widoki, historia szukania, presety, profile anonimizacji i wygląd na nowe urządzenie.",
    "Telefon: uchwyt panelu chowa się przy przewijaniu i edycji, żeby nie zasłaniał numerów wierszy.",
  ],
  en: [
    "Theme: System / Light / Dark — “System” follows your phone or computer setting.",
    "Excel-style questions: “Save / Don't save / Cancel” instead of the system OK/Cancel; view and preset names are typed in the app's own dialog.",
    "After paste, fill, “Replace all”, clear or moving cells — “Undo” right in the notification.",
    "Ctrl/⌘ + / opens a searchable shortcut list; toggle the theme quickly with Ctrl/⌘ + Alt + /.",
    "Reopening the same file offers to jump back to the sheet and cell where you left off.",
    "Settings backup (.json) in the ⋯ menu — views, search history, presets, anonymisation profiles and appearance for a new device.",
    "Phone: the panel handle hides while scrolling and editing so it doesn't cover row numbers.",
  ],
};
const SWB_SEEN_VERSION_KEY = "swb-seen-version";

function swbOfflineReady() {
  return !!(navigator.serviceWorker && navigator.serviceWorker.controller);
}

function swbSyncVersionItem() {
  const el = document.getElementById("appMenuBuild");
  if (!el) return;
  const v = typeof APP_BUILD_VERSION !== "undefined" ? APP_BUILD_VERSION : "";
  el.textContent = swbOfflineReady() ? t("menuVersionOffline", { v }) : t("menuVersion", { v });
  const kbd = document.getElementById("shortcutsMenuKbd");
  if (kbd && typeof swbKeyLabel === "function") kbd.textContent = `${swbKeyLabel("mod")} /`;
}

function swbShowWhatsNew() {
  const list = document.createElement("ul");
  list.className = "app-dialog-list";
  (SWB_WHATS_NEW[currentLang] || SWB_WHATS_NEW.pl).forEach((line) => list.append(Object.assign(document.createElement("li"), { textContent: line })));
  appDialog({
    title: t("whatsNewTitle", { v: APP_BUILD_VERSION }),
    content: list,
    actions: [{ label: t("close"), value: null, kind: "primary" }],
  });
}

async function swbCopyVersion() {
  const text = `Sheet Workbench ${APP_BUILD_VERSION}`;
  let ok = false;
  try { await navigator.clipboard.writeText(text); ok = true; } catch { /* brak uprawnień do schowka */ }
  toast(ok ? t("versionCopied", { v: APP_BUILD_VERSION }) : t("menuVersion", { v: APP_BUILD_VERSION }), "info", {
    action: t("whatsNewAction"), onAction: swbShowWhatsNew,
  });
}

// Po aktualizacji: jeden dymek „Zaktualizowano do vX” z „Co nowego” (pierwsze uruchomienie — cisza).
function swbAnnounceUpdate() {
  let prev = null;
  try { prev = localStorage.getItem(SWB_SEEN_VERSION_KEY); localStorage.setItem(SWB_SEEN_VERSION_KEY, APP_BUILD_VERSION); } catch { return; }
  if (!prev || prev === APP_BUILD_VERSION) return;
  setTimeout(() => toast(t("updatedTo", { v: APP_BUILD_VERSION }), "success", { action: t("whatsNewAction"), onAction: swbShowWhatsNew }), 1200);
}

// ── S13: kopia ustawień ──────────────────────────────────────────────────────
// Tylko preferencje i własne listy (widoki, historia, presety, profile) — NIGDY treść plików,
// szkice ani pomiary urządzenia. Nazwy kluczy są sprawdzane także przy wczytywaniu kopii.
const SWB_BACKUP_CATS = [
  { id: "views", label: "backupCatViews", keys: ["swb-saved-views-v1", "excel-workbench-qs-history", "excel-workbench-sort-presets", "swb-notes-search-v1", "swb-smart-cycle-cols"] },
  { id: "anon", label: "backupCatAnon", keys: ["swb-anon-profiles-v1", "swb-anon-prefs-v1"] },
  { id: "prefs", label: "backupCatPrefs", keys: null }, // reszta preferencji wyglądu i pracy
];
// pomiary/stany jednorazowe tego urządzenia i dane sesji — nie przenosimy
const SWB_BACKUP_SKIP = new Set(["swb-render-cost-v1", "swb-sessions-v1", "swb-scroll-nudge-v1", "swb-touch-range-hint-v1", SWB_SEEN_VERSION_KEY, "excel-workbench-cell-actions", "excel-workbench-cell-actions-collapsed"]);
const SWB_BACKUP_MAX_VALUE = 1024 * 1024;

function swbBackupKeyAllowed(key) {
  if (typeof key !== "string" || SWB_BACKUP_SKIP.has(key)) return false;
  return key.startsWith("swb-") || key.startsWith("excel-workbench-") || key === "localAi.prefs.v1";
}
function swbBackupCatOf(key) {
  const c = SWB_BACKUP_CATS.find((x) => x.keys && x.keys.includes(key));
  return c ? c.id : "prefs";
}
function swbBackupCollect() {
  const data = {};
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (!swbBackupKeyAllowed(k)) continue;
      const v = localStorage.getItem(k);
      if (v != null && v.length <= SWB_BACKUP_MAX_VALUE) data[k] = v;
    }
  } catch { /* brak dostępu do pamięci */ }
  return data;
}
function swbBackupCounts(data) {
  const n = { views: 0, anon: 0, prefs: 0 };
  Object.keys(data).forEach((k) => { n[swbBackupCatOf(k)]++; });
  return n;
}
function swbBackupChecks(counts) {
  const box = document.createElement("div");
  box.className = "app-dialog-checks";
  SWB_BACKUP_CATS.forEach((c) => {
    const label = document.createElement("label");
    label.className = "field checkbox";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = c.id;
    cb.checked = counts[c.id] > 0;
    cb.disabled = counts[c.id] === 0;
    const txt = document.createElement("span");
    txt.textContent = t(c.label);
    const small = Object.assign(document.createElement("small"), { textContent: t("backupCount", { n: counts[c.id] }) });
    label.append(cb, txt, small);
    box.append(label);
  });
  return box;
}
const swbCheckedCats = (box) => new Set(Array.from(box.querySelectorAll("input:checked")).map((cb) => cb.value));

async function swbOpenSettingsBackup() {
  const data = swbBackupCollect();
  const checks = swbBackupChecks(swbBackupCounts(data));
  const choice = await appDialog({
    title: t("backupTitle"),
    message: t("backupIntro"),
    content: checks,
    actions: [
      { label: t("backupSave"), value: "save", kind: "primary" },
      { label: t("backupLoad"), value: "load", kind: "secondary" },
      { label: t("dlgCancel"), value: null, kind: "secondary" },
    ],
  });
  if (choice === "save") {
    const cats = swbCheckedCats(checks);
    const out = {};
    Object.entries(data).forEach(([k, v]) => { if (cats.has(swbBackupCatOf(k))) out[k] = v; });
    if (!Object.keys(out).length) { toast(t("backupNothing"), "info"); return; }
    const payload = { app: "sheet-workbench", kind: "settings", version: APP_BUILD_VERSION, exportedAt: new Date().toISOString(), data: out };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `sheet-workbench-ustawienia-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    toast(t("backupSaved", { n: Object.keys(out).length }), "success");
  } else if (choice === "load") {
    swbPickSettingsFile();
  }
}

function swbPickSettingsFile() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".json,application/json";
  input.addEventListener("change", async () => {
    const f = input.files && input.files[0];
    if (f) await swbImportSettings(f);
  }, { once: true });
  input.click();
}

async function swbImportSettings(file) {
  let parsed = null;
  try { parsed = JSON.parse(await file.text()); } catch { /* nie JSON */ }
  if (!parsed || parsed.app !== "sheet-workbench" || parsed.kind !== "settings" || !parsed.data || typeof parsed.data !== "object") {
    toast(t("backupBadFile"), "error");
    return false;
  }
  const data = {};
  Object.entries(parsed.data).forEach(([k, v]) => { if (swbBackupKeyAllowed(k) && typeof v === "string" && v.length <= SWB_BACKUP_MAX_VALUE) data[k] = v; });
  const counts = swbBackupCounts(data);
  if (!Object.keys(data).length) { toast(t("backupNothing"), "info"); return false; }
  // podgląd przed wczytaniem: co jest w pliku, z którego dnia/wersji; wybór kategorii
  const checks = swbBackupChecks(counts);
  const when = parsed.exportedAt ? new Date(parsed.exportedAt) : null;
  const ok = await appDialog({
    title: t("backupImportTitle"),
    message: t("backupImportMsg", { when: when && !isNaN(when) ? when.toLocaleString(I18N[currentLang]?.locale || "pl-PL") : "—", v: String(parsed.version || "—") }),
    content: checks,
    actions: [
      { label: t("backupImportOk"), value: true, kind: "primary" },
      { label: t("dlgCancel"), value: null, kind: "secondary" },
    ],
  });
  if (ok !== true) return false;
  const cats = swbCheckedCats(checks);
  let n = 0;
  Object.entries(data).forEach(([k, v]) => {
    if (!cats.has(swbBackupCatOf(k))) return;
    try { localStorage.setItem(k, v); n++; } catch { /* brak miejsca */ }
  });
  // część ustawień (motyw, język, widoki) działa od razu, reszta czytana jest przy starcie
  if (typeof initTheme === "function") applyThemePref(getThemePref(), false);
  toast(t("backupLoaded", { n }), "success", {
    action: t("backupReload"),
    onAction: async () => { if (await confirmReloadWithUnsaved()) location.reload(); },
  });
  return true;
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("appMenuVersion")?.addEventListener("click", () => {
    document.getElementById("appMenu").hidden = true;
    document.getElementById("appMenuBtn")?.setAttribute("aria-expanded", "false");
    swbCopyVersion();
  });
  document.getElementById("settingsBackupItem")?.addEventListener("click", () => {
    document.getElementById("appMenu").hidden = true;
    document.getElementById("appMenuBtn")?.setAttribute("aria-expanded", "false");
    swbOpenSettingsBackup();
  });
  swbSyncVersionItem();
  swbAnnounceUpdate();
});
