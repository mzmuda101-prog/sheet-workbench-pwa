// recent-files.js — „Ostatnio otwierane” na ekranie startowym (audyt QoL 2026-11-06, ekran startowy).
//
// Safari (także aplikacja w Docku i iPhone) nie daje stronie trwałego dostępu do pliku na
// dysku, więc trzymamy LOKALNĄ KOPIĘ ostatnich plików w IndexedDB (tylko na tym urządzeniu,
// nic nie wychodzi z komputera). W Chrome/Edge zapamiętujemy też uchwyt pliku: klik otwiera
// wtedy aktualną wersję z dysku i „Zapisz” dalej nadpisuje oryginał; bez zgody albo gdy
// pliku już tam nie ma — otwiera się kopia.
//
// Po zapisie (w miejscu albo pobraniem) kopia na liście dostaje zapisaną wersję — inaczej
// w Safari „ostatni” plik otwierałby się sprzed zmian. Przykładowego pliku nie zapamiętujemy.
// Kopia ustawień (.json, app-settings.js) tej listy NIE obejmuje — to treść plików.

const RECENT_DB = "swb-recent";
const RECENT_META = "meta";
const RECENT_BYTES = "bytes";
const RECENT_MAX = 5;
const RECENT_MAX_BYTES = 30 * 1024 * 1024; // większy plik: bez kopii (w Chrome zostaje sam uchwyt)

const swbRecent = (() => {
  let currentId = "";
  let skipNext = false;

  // ── IndexedDB ──────────────────────────────────────────────────────────────
  let dbPromise = null;
  function db() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        if (typeof indexedDB === "undefined") { reject(new Error("no indexedDB")); return; }
        const req = indexedDB.open(RECENT_DB, 1);
        req.onupgradeneeded = () => {
          req.result.createObjectStore(RECENT_META, { keyPath: "id" });
          req.result.createObjectStore(RECENT_BYTES);
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
      const t = d.transaction([RECENT_META, RECENT_BYTES], mode);
      let result;
      Promise.resolve(fn(t.objectStore(RECENT_META), t.objectStore(RECENT_BYTES))).then((r) => { result = r; });
      t.oncomplete = () => resolve(result);
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error);
    });
  }
  const reqP = (r) => new Promise((resolve, reject) => { r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error); });
  const list = () => tx("readonly", (s) => reqP(s.getAll()))
    .then((recs) => (recs || []).sort((a, b) => (b.openedAt || 0) - (a.openedAt || 0)))
    .catch(() => []);
  const getBytes = (id) => tx("readonly", (s, b) => reqP(b.get(id))).catch(() => null);
  const remove = (id) => tx("readwrite", (s, b) => { s.delete(id); b.delete(id); }).catch(() => {});

  const keyOf = (name, size, lastModified) => `${name}|${size}|${lastModified}`;
  const newId = () => `r${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;

  // Wpis na górę listy; ten sam plik (nazwa + rozmiar + data) nie dubluje się. Nadmiar precz.
  async function put(entry, bytes) {
    const recs = await list();
    const same = recs.find((r) => r.id === entry.id || r.key === entry.key);
    if (same) entry.id = same.id;
    const drop = recs.filter((r) => r.id !== entry.id).slice(RECENT_MAX - 1).map((r) => r.id);
    await tx("readwrite", (s, b) => {
      s.put(entry);
      if (bytes) b.put(bytes, entry.id);
      else b.delete(entry.id);
      drop.forEach((id) => { s.delete(id); b.delete(id); });
    });
    return entry.id;
  }

  async function record(file, handle) {
    const size = file.size || 0;
    const copy = size <= RECENT_MAX_BYTES;
    if (!copy && !handle) return; // nie będzie czym otworzyć
    const bytes = copy ? new Uint8Array(await file.arrayBuffer()) : null;
    const entry = {
      id: newId(),
      key: keyOf(file.name, size, file.lastModified),
      name: file.name,
      size,
      lastModified: file.lastModified || Date.now(),
      openedAt: Date.now(),
      handle: handle || null,
      hasCopy: copy,
    };
    currentId = await put(entry, bytes);
  }

  // Zapis (saveWorkbook / „Zapisz jako…”): kopia na liście = zapisana wersja. Pobranie w
  // Safari to inny plik niż oryginał, więc stary uchwyt (Chrome) wtedy przestaje pasować.
  async function noteSaved(name, data, handle) {
    if (!currentId || !data) return;
    try {
      const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
      const recs = await list();
      const prev = recs.find((r) => r.id === currentId);
      if (!prev) return;
      const now = Date.now();
      const copy = bytes.length <= RECENT_MAX_BYTES;
      await put({
        ...prev,
        key: keyOf(name || prev.name, bytes.length, now),
        name: name || prev.name,
        size: bytes.length,
        lastModified: now,
        openedAt: now,
        handle: handle || null,
        hasCopy: copy,
      }, copy ? bytes : null);
    } catch (_) { /* lista to wygoda, nie warunek zapisu */ }
  }

  // ── otwieranie z listy ─────────────────────────────────────────────────────
  async function open(rec) {
    if (typeof ensureXlsxLibs === "function" && !(await ensureXlsxLibs(true))) return false;
    let file = null;
    let handle = null;
    if (rec.handle && typeof rec.handle.requestPermission === "function") {
      try {
        if ((await rec.handle.requestPermission({ mode: "readwrite" })) === "granted") {
          file = await rec.handle.getFile();
          handle = rec.handle;
        }
      } catch (_) { file = null; handle = null; }
      if (!file && rec.hasCopy) toast(t("recentFromCopy", { name: rec.name }), "info");
    }
    if (!file) {
      const bytes = rec.hasCopy ? await getBytes(rec.id) : null;
      if (!bytes) {
        toast(t("recentGone", { name: rec.name }), "warning");
        await remove(rec.id);
        render();
        return false;
      }
      file = new File([bytes], rec.name, { type: typeof XLSX_MIME !== "undefined" ? XLSX_MIME : "", lastModified: rec.lastModified });
    }
    const ok = await handleFile(file, handle);
    if (ok && typeof openDefaultSheet === "function") openDefaultSheet();
    return ok;
  }

  // ── lista na ekranie startowym ─────────────────────────────────────────────
  function fmtSize(n) {
    const loc = (typeof I18N !== "undefined" && I18N[currentLang]?.locale) || "pl-PL";
    if (n < 1024) return `${n} B`;
    if (n < 1024 * 1024) return `${Math.round(n / 1024)} KB`;
    return `${(n / 1024 / 1024).toLocaleString(loc, { maximumFractionDigits: 1 })} MB`;
  }
  function when(ts) {
    const d = new Date(ts);
    const loc = (typeof I18N !== "undefined" && I18N[currentLang]?.locale) || "pl-PL";
    const time = d.toLocaleTimeString(loc, { hour: "2-digit", minute: "2-digit" });
    const y = new Date(); y.setDate(y.getDate() - 1);
    if (d.toDateString() === new Date().toDateString()) return t("draftToday", { time });
    if (d.toDateString() === y.toDateString()) return t("draftYesterday", { time });
    return `${d.toLocaleDateString(loc, { day: "numeric", month: "long" })}, ${time}`;
  }

  let renderJob = 0;
  async function render() {
    const job = ++renderJob;
    const host = document.getElementById("recentFiles");
    if (!host) return;
    if (workbook) { host.replaceChildren(); host.hidden = true; return; }
    const recs = await list();
    if (job !== renderJob || workbook) return;
    host.replaceChildren();
    host.hidden = !recs.length;
    if (!recs.length) return;
    const head = Object.assign(document.createElement("div"), { className: "recent-head" });
    const cap = Object.assign(document.createElement("div"), { className: "recent-cap", id: "recentCap", textContent: t("recentTitle") });
    const clear = Object.assign(document.createElement("button"), { type: "button", className: "recent-clear", textContent: t("recentClear") });
    clear.addEventListener("click", async () => {
      const ok = await appConfirm(t("recentClearAsk"), { title: t("recentClearTitle"), okLabel: t("recentClear"), danger: true });
      if (!ok) return;
      await tx("readwrite", (s, b) => { s.clear(); b.clear(); }).catch(() => {});
      currentId = "";
      render();
    });
    head.append(cap, clear);
    const ul = Object.assign(document.createElement("ul"), { className: "recent-list" });
    ul.setAttribute("aria-labelledby", "recentCap");
    recs.forEach((r) => {
      const li = Object.assign(document.createElement("li"), { className: "recent-item" });
      const btn = Object.assign(document.createElement("button"), { type: "button", className: "recent-open" });
      btn.append(
        Object.assign(document.createElement("span"), { className: "recent-name", textContent: r.name }),
        Object.assign(document.createElement("span"), { className: "recent-meta", textContent: `${when(r.openedAt)} · ${fmtSize(r.size)}` }),
      );
      btn.addEventListener("click", async () => {
        btn.disabled = true;
        try { await open(r); } catch (e) { log(String(e?.message || e), "error"); toast(t("recentOpenFailed"), "error"); }
        btn.disabled = false;
      });
      const del = Object.assign(document.createElement("button"), { type: "button", className: "recent-remove", textContent: "✕" });
      del.setAttribute("aria-label", t("recentRemove", { name: r.name }));
      del.title = t("recentRemove", { name: r.name });
      del.addEventListener("click", async () => {
        await remove(r.id);
        if (r.id === currentId) currentId = "";
        render();
      });
      li.append(btn, del);
      ul.append(li);
    });
    host.append(head, ul, Object.assign(document.createElement("p"), { className: "recent-note", textContent: t("recentNote") }));
  }

  // ── podpięcie pod aplikację ────────────────────────────────────────────────
  function init() {
    if (typeof window.handleFile !== "function") return;
    const orig = window.handleFile;
    window.handleFile = async function handleFileRecent(file, handle, ...rest) {
      const skip = skipNext;
      skipNext = false;
      const r = await orig.call(this, file, handle, ...rest);
      if (r && workbook && file && currentFileName === file.name) {
        currentId = ""; // zapis poprzedniego pliku nie może trafić we wpis nowego
        if (!skip) record(file, handle).catch(() => {}).then(render);
      }
      render();
      return r;
    };
    render();
  }

  return { init, render, noteSaved, skipNextOpen: () => { skipNext = true; }, _list: list, _current: () => currentId };
})();

// po drafts.js i session-resume.js (te też owijają handleFile przy DOMContentLoaded)
document.addEventListener("DOMContentLoaded", () => swbRecent.init());
