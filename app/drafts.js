// drafts.js — automatyczny szkic niezapisanej pracy + odzyskiwanie po zamknięciu aplikacji.
// Przeniesione z Documents Workbench (2026-10-01), ten sam model, inne „co zapisać”.
//
// iOS / Android potrafią zamknąć aplikację w tle (brak pamięci) bez żadnego zdarzenia
// „zamykam się”; na komputerze — awaria przeglądarki, prąd, restart. Wtedy niezapisane
// zmiany przepadały. Tu: szkic w IndexedDB (tylko na tym urządzeniu), przy następnym
// otwarciu karta „Niezapisana praca … [Przywróć] [Odrzuć]” na ekranie startowym.
//
// Co i kiedy:
//   - bajty pliku (originalFileBytes) — tylko gdy się zmieniły (otwarcie pliku),
//   - zmienione komórki (pendingEdits — to samo, co trafia do zapisu) + otwarty arkusz —
//     małe, ~1,5 s po zmianie i NATYCHMIAST przy przejściu w tło (visibilitychange/pagehide).
// Szkic znika: po zapisie, po „Odrzuć”, po otwarciu innego pliku (aplikacja pyta o porzucenie
// zmian), gdy nie ma już żadnej zmienionej komórki, i sam po 7 dniach. Każde okno/karta ma
// własny szkic; żyjące okna odpowiadają na BroadcastChannel i nie są pokazywane jako zgubiona
// praca. Uchwyt pliku (Chrome/Edge) też trafia do szkicu — chyba że oryginał na dysku zmienił
// się od czasu szkicu (wtedy „Zapisz” utworzy kopię).

const DRAFT_DB = "swb-drafts";
const DRAFT_STORE = "drafts"; // opis + zmienione komórki (małe, zapisywane często)
const DRAFT_BYTES = "bytes"; // bajty pliku (duże, zapisywane tylko gdy się zmieniły)
const DRAFT_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000; // prośba Mateusza: 7 dni
const DRAFT_DEBOUNCE_MS = 1500;
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const swbDrafts = (() => {
  let sessionId = newDraftId();
  let timer = 0;
  let saving = null;
  let storedBytes = null;
  let stamp = null; // { size, lastModified } oryginału na dysku, gdy szkic powstał
  let persistAsked = false;
  let suspended = 0; // >0 = trwa przywracanie (handleFile nie kasuje szkicu)
  const channel = typeof BroadcastChannel === "function" ? new BroadcastChannel("swb-drafts") : null;

  function newDraftId() {
    return `s${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  }
  function editCount(edits) {
    return Object.values(edits || {}).reduce((n, cells) => n + Object.keys(cells || {}).length, 0);
  }

  // ── IndexedDB ──────────────────────────────────────────────────────────────
  let dbPromise = null;
  function db() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        if (typeof indexedDB === "undefined") { reject(new Error("no indexedDB")); return; }
        const req = indexedDB.open(DRAFT_DB, 1);
        req.onupgradeneeded = () => {
          req.result.createObjectStore(DRAFT_STORE, { keyPath: "id" });
          req.result.createObjectStore(DRAFT_BYTES);
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      }).catch((e) => { dbPromise = null; throw e; });
    }
    return dbPromise;
  }
  async function tx(mode, fn) {
    const d = await db();
    return new Promise((resolve, reject) => {
      const t = d.transaction([DRAFT_STORE, DRAFT_BYTES], mode);
      let result;
      Promise.resolve(fn(t.objectStore(DRAFT_STORE), t.objectStore(DRAFT_BYTES))).then((r) => { result = r; });
      t.oncomplete = () => resolve(result);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    });
  }
  const reqP = (r) => new Promise((resolve, reject) => { r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error); });
  const getRec = (id) => tx("readonly", async (s, b) => {
    const rec = await reqP(s.get(id));
    if (rec) rec.bytes = await reqP(b.get(id));
    return rec;
  });
  const allRecs = () => tx("readonly", (s) => reqP(s.getAll()));
  const delRec = (id) => tx("readwrite", (s, b) => { s.delete(id); b.delete(id); }).catch(() => {});

  // ── zapis szkicu ───────────────────────────────────────────────────────────
  async function writeNow() {
    clearTimeout(timer);
    timer = 0;
    if (!originalFileBytes || !workbook || !hasUnsavedChanges) return;
    const count = editCount(pendingEdits);
    if (!count) { await delRec(sessionId); storedBytes = null; return; } // np. cofnięte do stanu z pliku
    const bytes = originalFileBytes;
    const id = sessionId;
    if (!stamp && currentFileHandle) {
      try { const f = await currentFileHandle.getFile(); stamp = { size: f.size, lastModified: f.lastModified }; } catch (_) { /* brak zgody — bez odcisku */ }
    }
    const rec = {
      id, fileName: currentFileName || "arkusz.xlsx", savedAt: Date.now(), stamp,
      // + wiersz nagłówka: bez tego arkusz po odzyskaniu wczytałby się z nagłówkiem w 1. wierszu
      data: {
        edits: pendingEdits, sheet: currentSheetName || "", changes: count,
        headerRow: headerRowEl?.value || "", autoHeader: !!autoHeaderRowEl?.checked,
      },
      handle: currentFileHandle || null,
    };
    const writeBytes = storedBytes !== bytes;
    try {
      await tx("readwrite", (s, b) => {
        if (writeBytes) b.put(bytes, id);
        try { s.put(rec); } catch (_) { s.put({ ...rec, handle: null }); } // uchwytu nie da się zapisać — bez niego
      });
      storedBytes = bytes;
      if (!persistAsked) {
        persistAsked = true;
        navigator.storage?.persisted?.().then((p) => { if (!p) navigator.storage.persist?.(); }).catch(() => {});
      }
    } catch (e) {
      if (typeof log === "function") log(`Szkic niezapisany: ${e?.message || e}`, "warning");
    }
  }

  function save(delay = DRAFT_DEBOUNCE_MS) {
    if (suspended) return;
    clearTimeout(timer);
    timer = setTimeout(() => { saving = writeNow().finally(() => { saving = null; }); }, delay);
  }
  function saveImmediately() {
    if (suspended || !hasUnsavedChanges || !originalFileBytes) return;
    saving = writeNow().finally(() => { saving = null; });
  }
  async function dropOwn() {
    clearTimeout(timer);
    timer = 0;
    if (saving) await saving.catch(() => {});
    storedBytes = null;
    stamp = null;
    await delRec(sessionId);
  }

  // ── odzyskiwanie ───────────────────────────────────────────────────────────
  function aliveIds(timeoutMs = 250) {
    if (!channel) return Promise.resolve(new Set());
    const ids = new Set();
    const onMsg = (e) => { if (e.data?.type === "alive" && e.data.id) ids.add(e.data.id); };
    channel.addEventListener("message", onMsg);
    channel.postMessage({ type: "who" });
    return new Promise((r) => setTimeout(() => { channel.removeEventListener("message", onMsg); r(ids); }, timeoutMs));
  }
  channel?.addEventListener("message", (e) => {
    if (e.data?.type === "who" && workbook) channel.postMessage({ type: "alive", id: sessionId });
  });

  async function recoverable() {
    let recs;
    try { recs = await allRecs(); } catch (_) { return []; }
    const now = Date.now();
    const expired = recs.filter((r) => !r.savedAt || now - r.savedAt > DRAFT_MAX_AGE_MS);
    expired.forEach((r) => delRec(r.id));
    const alive = await aliveIds();
    return recs
      .filter((r) => !expired.includes(r) && r.id !== sessionId && !alive.has(r.id))
      .sort((a, b) => b.savedAt - a.savedAt);
  }

  // Zmienione komórki z powrotem do skoroszytu — tak samo jak updateSheetCell (styl i format zostają).
  function applyEdits(edits) {
    Object.entries(edits || {}).forEach(([sheetName, cells]) => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) return;
      Object.entries(cells || {}).forEach(([ref, p]) => {
        if (p === null) { delete sheet[ref]; return; }
        const prev = sheet[ref];
        const cell = { v: p.v, t: p.t };
        if (prev && prev.s) cell.s = prev.s;
        if (prev && prev.z) cell.z = prev.z;
        sheet[ref] = cell;
      });
    });
    // to, co pójdzie do zapisu — obiekty prosto ze szkicu (IndexedDB zachowuje daty)
    pendingEdits = {};
    Object.entries(edits || {}).forEach(([sn, cells]) => { pendingEdits[sn] = { ...cells }; });
    if (typeof bumpSheetDataStamp === "function") bumpSheetDataStamp();
  }

  async function restore(id) {
    const rec = await getRec(id);
    if (!rec?.bytes) { toast(t("draftGone"), "info"); renderRecovery(); return false; }
    if (typeof ensureXlsxLibs === "function" && !(await ensureXlsxLibs(true))) return false;
    let handle = null;
    if (rec.handle && typeof rec.handle.requestPermission === "function") {
      try {
        const perm = await rec.handle.requestPermission({ mode: "readwrite" });
        if (perm === "granted") {
          const f = await rec.handle.getFile();
          if (rec.stamp && (f.size !== rec.stamp.size || f.lastModified !== rec.stamp.lastModified)) toast(t("draftOriginalChanged", { name: rec.fileName }), "warning");
          else handle = rec.handle;
        }
      } catch (_) { handle = null; }
    }
    const bytes = rec.bytes instanceof Uint8Array ? rec.bytes : new Uint8Array(rec.bytes);
    suspended++;
    try {
      await handleFile(new File([bytes], rec.fileName, { type: XLSX_MIME }), handle);
      if (!workbook) return false;
      applyEdits(rec.data?.edits);
    } finally { suspended--; }
    sessionId = rec.id;
    storedBytes = null;
    stamp = handle ? rec.stamp : null;
    setDirtyState(true);
    // otwarty wtedy arkusz — od razu z powrotem w tabeli
    const sheetName = rec.data?.sheet;
    if (sheetName && workbook.Sheets[sheetName] && sheetSelect) {
      sheetSelect.value = sheetName;
      if (autoHeaderRowEl && typeof rec.data.autoHeader === "boolean") autoHeaderRowEl.checked = rec.data.autoHeader;
      if (headerRowEl && rec.data.headerRow) headerRowEl.value = rec.data.headerRow;
      loadBtn?.click();
    }
    toast(t("draftRestored", { name: rec.fileName }), "success");
    saveImmediately();
    renderRecovery();
    return true;
  }

  async function discard(id, name) {
    if (!window.confirm(t("draftDiscardConfirm", { name }))) return;
    await delRec(id);
    renderRecovery();
  }

  // ── karta na ekranie startowym ─────────────────────────────────────────────
  function when(ts) {
    const d = new Date(ts);
    const loc = (typeof I18N !== "undefined" && I18N[currentLang]?.locale) || "pl-PL";
    const time = d.toLocaleTimeString(loc, { hour: "2-digit", minute: "2-digit" });
    const today = new Date();
    const y = new Date(); y.setDate(today.getDate() - 1);
    if (d.toDateString() === today.toDateString()) return t("draftToday", { time });
    if (d.toDateString() === y.toDateString()) return t("draftYesterday", { time });
    return `${d.toLocaleDateString(loc, { day: "numeric", month: "long" })}, ${time}`;
  }

  let renderJob = 0;
  async function renderRecovery() {
    const job = ++renderJob;
    const host = document.getElementById("draftRecovery");
    if (!host) return;
    if (workbook) { host.replaceChildren(); host.hidden = true; return; }
    const recs = await recoverable();
    if (job !== renderJob || workbook) return;
    host.replaceChildren();
    host.hidden = !recs.length;
    recs.slice(0, 5).forEach((r) => {
      const card = document.createElement("div");
      card.className = "draft-card";
      card.setAttribute("role", "group");
      const head = Object.assign(document.createElement("div"), { className: "draft-head", textContent: t("draftTitle") });
      const name = Object.assign(document.createElement("div"), { className: "draft-name", textContent: r.fileName });
      const n = r.data?.changes || 0;
      const meta = Object.assign(document.createElement("div"), {
        className: "draft-meta",
        textContent: [t("draftChangedAt", { when: when(r.savedAt) }), n ? t("draftChanges", { count: n }) : "", r.data?.sheet ? t("draftSheet", { sheet: r.data.sheet }) : ""].filter(Boolean).join(" · "),
      });
      const actions = Object.assign(document.createElement("div"), { className: "draft-actions" });
      const ok = Object.assign(document.createElement("button"), { type: "button", className: "btn primary", textContent: t("draftRestore") });
      ok.addEventListener("click", async () => {
        ok.disabled = true;
        try { await restore(r.id); } catch (e) { log(String(e?.message || e), "error"); toast(t("draftRestoreFailed"), "error"); ok.disabled = false; }
      });
      const no = Object.assign(document.createElement("button"), { type: "button", className: "btn ghost", textContent: t("draftDiscard") }); // w Sheet „btn” jest wypełniony
      no.addEventListener("click", () => discard(r.id, r.fileName));
      actions.append(ok, no);
      card.append(head, name, meta, actions);
      host.append(card);
    });
    if (recs.length) host.append(Object.assign(document.createElement("p"), { className: "draft-note", textContent: t("draftNote") }));
  }

  // ── podpięcie pod aplikację ────────────────────────────────────────────────
  function init() {
    const origSetDirty = window.setDirtyState;
    window.setDirtyState = function setDirtyStateDrafts(isDirty, ...rest) {
      const r = origSetDirty.call(this, isDirty, ...rest);
      if (isDirty) save();
      else if (!suspended) dropOwn(); // zapisane / nowy plik
      return r;
    };
    // nowy plik: aplikacja już zapytała o porzucenie zmian → szkic tej sesji precz
    const origHandleFile = window.handleFile;
    window.handleFile = async function handleFileDrafts(...args) {
      if (!suspended) {
        if (!confirmDiscardChanges()) return;
        await dropOwn();
        sessionId = newDraftId();
      }
      const r = await origHandleFile.apply(this, args);
      renderRecovery();
      return r;
    };
    // ostatnia chwila przed ewentualnym zabiciem karty przez system
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") saveImmediately(); });
    window.addEventListener("pagehide", saveImmediately);
    renderRecovery();
  }

  return { init, save, saveImmediately, renderRecovery, restore, recoverable, _id: () => sessionId, _flush: () => (saving || Promise.resolve()) };
})();

// Inny plik przy niezapisanych zmianach: pytamy (dawniej Sheet podmieniał plik bez słowa —
// zmienione komórki przepadały; Documents Workbench pytał od zawsze).
function confirmDiscardChanges() {
  return !workbook || !hasUnsavedChanges || window.confirm(t("openDiscardWarn"));
}

document.addEventListener("DOMContentLoaded", () => swbDrafts.init());
