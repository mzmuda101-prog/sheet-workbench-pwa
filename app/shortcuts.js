// shortcuts.js — JEDEN rejestr skrótów klawiszowych + okno „Skróty i pomoc” (Ctrl/⌘+/).
//
// Kontrakt UX z Documents Workbench (2026-11-07): lista skrótów nie może rozjeżdżać się
// z kodem, więc i ściąga w panelu „Pomoc”, i okno pod Ctrl/⌘+/ rysują się z tej samej tablicy
// SWB_SHORTCUTS. Dopisujesz nowy skrót → dopisz go tutaj (test shortcuts-palette pilnuje,
// żeby każdy wpis miał tekst w obu językach i żeby Ctrl/⌘+/ otwierało okno).
//
// Klawisze: "mod" = ⌘ na Macu/iPadzie/iPhonie, Ctrl gdzie indziej; "ctrl" = zawsze Ctrl
// (Ctrl+D/R — jak w Excelu, ⌘R przeładowałby stronę); "alt" = ⌥ na Macu, Alt gdzie indziej;
// "@klucz" = przetłumaczony napis (np. „Strzałki”). Kombinacje alternatywne: tablica tablic.
// Litery nigdy bez modyfikatora (pisanie na klawiaturze tabletu), F1–F12 tylko jako alias.

const SWB_SHORTCUT_CATS = [
  { id: "search", label: "scCatSearch" },
  { id: "table", label: "scCatTable" },
  { id: "edit", label: "scCatEdit" },
  { id: "file", label: "scCatFile" },
  { id: "app", label: "scCatApp" },
];

const SWB_SHORTCUTS = [
  // szukanie
  { cat: "search", keys: ["mod", "shift", "F"], text: "quickSearch" },
  { cat: "search", keys: ["/"], text: "quickSearchSlash" },
  { cat: "search", keys: ["Enter"], text: "quickSearchEnter" },
  { cat: "search", keys: [["↓"], ["↑"]], text: "quickSearchArrows" },
  { cat: "search", keys: ["@spaceKey"], text: "qsSpaceToggle" },
  { cat: "search", keys: ["Tab"], text: "qsTabCycle" },
  { cat: "search", keys: ["Esc"], text: "qsEscBack" },
  { cat: "search", keys: ["mod", "Enter"], text: "filter" },
  { cat: "search", keys: ["mod", "shift", "X"], text: "resetFilters" },
  // tabela i zaznaczenie
  { cat: "table", keys: ["@clickKey"], text: "clickCell" },
  { cat: "table", keys: ["@arrows"], text: "arrowFocus" },
  { cat: "table", keys: [["←"], ["→"]], text: "arrowRowScroll" },
  { cat: "table", keys: ["shift", "@spaceKey"], text: "toggleSelectionKind" },
  { cat: "table", keys: ["shift", "@clickKey"], text: "showCellSelection" },
  { cat: "table", keys: ["shift", "@arrows"], text: "moveCellSelection" },
  { cat: "table", keys: ["Esc"], text: "clearRowFocus" },
  { cat: "table", keys: ["shift", "Esc"], text: "clearCellSelection" },
  { cat: "table", keys: ["mod", "alt", "R"], text: "resetSortShortcut" },
  { cat: "table", keys: ["mod", "alt", "W"], text: "resetWidths" },
  { cat: "table", keys: ["mod", "alt", "K"], text: "columns" },
  { cat: "table", keys: ["mod", "alt", "1 2 3"], text: "regionJump" },
  { cat: "table", keys: ["F6"], text: "regionCycle" },
  { cat: "table", keys: ["Tab"], text: "gridTabEnter" },
  // edycja
  { cat: "edit", keys: ["F2"], text: "editOpen" },
  { cat: "edit", keys: [["Enter"], ["shift", "Enter"]], text: "editCommitVertical" },
  { cat: "edit", keys: [["Tab"], ["shift", "Tab"]], text: "editCommitHorizontal" },
  { cat: "edit", keys: ["shift", "@arrows"], text: "editMoveArrows" },
  { cat: "edit", keys: [["←"], ["→"]], text: "editCaret" },
  { cat: "edit", keys: ["Esc"], text: "editCancel" },
  { cat: "edit", keys: ["mod", "Z"], text: "undoShortcut" },
  { cat: "edit", keys: [["mod", "shift", "Z"], ["ctrl", "Y"]], text: "redoShortcut" },
  { cat: "edit", keys: ["mod", "V"], text: "pasteShortcut" },
  { cat: "edit", keys: ["mod", "shift", "V"], text: "pasteValuesShortcut" },
  { cat: "edit", keys: ["ctrl", "D"], text: "fillDownShortcut" },
  { cat: "edit", keys: ["ctrl", "R"], text: "fillRightShortcut" },
  { cat: "edit", keys: ["⠿"], text: "fillHandleShortcut" },
  { cat: "edit", keys: ["shift", "F2"], text: "noteEditShortcut" },
  { cat: "edit", keys: ["mod", "alt", "N"], text: "noteNextShortcut" },
  { cat: "edit", keys: ["#note"], text: "noteMarkerInfo" },
  // plik
  { cat: "file", keys: ["mod", "S"], text: "saveInPlace" },
  { cat: "file", keys: ["mod", "alt", "S"], text: "saveAs" },
  { cat: "file", keys: ["mod", "shift", "E"], text: "exportCsv" },
  // aplikacja
  { cat: "app", keys: ["mod", "/"], text: "shortcutsPaletteShortcut" },
  { cat: "app", keys: ["mod", "alt", "/"], text: "themeQuickShortcut" },
  { cat: "app", keys: ["Esc"], text: "closeModalPanel" },
];

const SWB_IS_APPLE = (() => {
  try {
    const p = (navigator.userAgentData && navigator.userAgentData.platform) || navigator.platform || "";
    return /mac|iphone|ipad|ipod/i.test(p) || (/Macintosh/.test(navigator.userAgent) && navigator.maxTouchPoints > 1);
  } catch { return false; }
})();

function swbKeyLabel(k) {
  if (k === "mod") return SWB_IS_APPLE ? "⌘" : "Ctrl";
  if (k === "ctrl") return "Ctrl";
  if (k === "alt") return SWB_IS_APPLE ? "⌥" : "Alt";
  if (k === "shift") return SWB_IS_APPLE ? "⇧" : "Shift";
  if (k.startsWith("@")) return t(k.slice(1));
  return k;
}

// Kombinacja jako tekst (do wyszukiwania i dla czytnika): „⌘ + ⇧ + F”.
function swbShortcutKeysText(sc) {
  const combos = Array.isArray(sc.keys[0]) ? sc.keys : [sc.keys];
  return combos.map((c) => c.map(swbKeyLabel).join(" + ")).join(" / ");
}

function swbShortcutRow(sc) {
  const row = document.createElement("div");
  row.className = "shortcut-row";
  const combos = Array.isArray(sc.keys[0]) ? sc.keys : [sc.keys];
  combos.forEach((combo, ci) => {
    if (ci) row.append(Object.assign(document.createElement("span"), { textContent: "/" }));
    combo.forEach((k, ki) => {
      if (ki) row.append(Object.assign(document.createElement("span"), { textContent: "+" }));
      const kbd = document.createElement("kbd");
      if (k === "#note") { kbd.className = "shortcut-note-kbd"; kbd.setAttribute("aria-hidden", "true"); }
      else kbd.textContent = swbKeyLabel(k);
      row.append(kbd);
    });
  });
  const span = document.createElement("span");
  span.textContent = t(sc.text);
  span.dataset.shortcutText = sc.text; // setShortcutTexts (language.js) i testy szukają wpisu po kluczu
  row.append(span);
  return row;
}

// Ściąga w panelu „Pomoc” (sidebar) — te same wpisy, pogrupowane jak w oknie.
function swbRenderShortcutPanel() {
  const host = document.getElementById("shortcutRegistryList");
  if (!host) return;
  host.replaceChildren(...SWB_SHORTCUT_CATS.map((cat) => {
    const block = document.createElement("div");
    block.className = "shortcut-hint";
    block.append(Object.assign(document.createElement("div"), { className: "shortcut-section-title", textContent: t(cat.label) }));
    SWB_SHORTCUTS.filter((s) => s.cat === cat.id).forEach((s) => block.append(swbShortcutRow(s)));
    return block;
  }));
}

// Okno „Skróty i pomoc” z wyszukiwarką (Ctrl/⌘+/ albo menu ⋯ → Skróty klawiszowe).
let swbShortcutsOpen = false;
function swbOpenShortcuts() {
  if (swbShortcutsOpen) return;
  swbShortcutsOpen = true;
  const wrap = document.createElement("div");
  wrap.className = "sc-palette";
  const search = document.createElement("input");
  search.type = "search";
  search.className = "modal-search sc-search";
  search.placeholder = t("scSearchPlaceholder");
  search.setAttribute("aria-label", t("scSearchPlaceholder"));
  search.autocomplete = "off";
  search.spellcheck = false;
  const list = document.createElement("div");
  list.className = "sc-list";
  const empty = Object.assign(document.createElement("p"), { className: "sc-empty", textContent: t("scNoMatch") });
  const render = () => {
    const q = search.value.trim().toLowerCase();
    const blocks = [];
    SWB_SHORTCUT_CATS.forEach((cat) => {
      const items = SWB_SHORTCUTS.filter((s) => s.cat === cat.id
        && (!q || `${t(s.text)} ${swbShortcutKeysText(s)} ${t(cat.label)}`.toLowerCase().includes(q)));
      if (!items.length) return;
      const block = document.createElement("section");
      block.className = "shortcut-hint";
      block.append(Object.assign(document.createElement("h3"), { className: "shortcut-section-title", textContent: t(cat.label) }));
      items.forEach((s) => block.append(swbShortcutRow(s)));
      blocks.push(block);
    });
    list.replaceChildren(...(blocks.length ? blocks : [empty]));
  };
  search.addEventListener("input", render);
  render();
  const more = document.createElement("p");
  more.className = "sc-more";
  more.textContent = t("scMoreInPanel");
  wrap.append(search, list, more);
  appDialog({
    title: t("scTitle"),
    content: wrap,
    className: "sc-dialog",
    input: null,
    actions: [
      { label: t("scOpenHelpPanel"), value: "panel", kind: "secondary" },
      { label: t("close"), value: null, kind: "primary" },
    ],
  }).then((v) => {
    swbShortcutsOpen = false;
    if (v === "panel") swbOpenHelpPanel();
  });
  // z klawiatury od razu do wyszukiwania (dotyk: bez fokusu — klawiatura ekranowa by wyskoczyła)
  if (appDialogWantsFocus()) requestAnimationFrame(() => requestAnimationFrame(() => search.focus()));
}

// Pełna ściąga (operatory szukania, scalanie grup) zostaje w panelu „Pomoc”.
function swbOpenHelpPanel() {
  const panel = document.getElementById("panel-shortcuts");
  if (typeof setSidebarOpen === "function") setSidebarOpen(true);
  if (panel) {
    panel.open = true;
    requestAnimationFrame(() => panel.scrollIntoView({ block: "start", behavior: "smooth" }));
  }
}

document.addEventListener("DOMContentLoaded", () => {
  swbRenderShortcutPanel();
  const item = document.getElementById("shortcutsMenuItem");
  if (item) item.addEventListener("click", () => {
    const menu = document.getElementById("appMenu");
    if (menu) menu.hidden = true;
    document.getElementById("appMenuBtn")?.setAttribute("aria-expanded", "false");
    swbOpenShortcuts();
  });
});
