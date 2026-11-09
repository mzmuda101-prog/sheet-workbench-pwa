// app-dialog.js — własne okna zamiast systemowych confirm() / prompt() (2026-11-07).
//
// Wspólny kontrakt UX z Documents Workbench (wzorzec: Smart Kalkulator). Systemowe
// „OK / Anuluj” ukrywało konsekwencję (szczególnie na telefonie: „OK” = co właściwie?),
// a prompt() na iOS wyskakiwał nad klawiaturą bez opisu. Tu jedno okno z tytułem,
// opisem i KONKRETNYMI przyciskami („Zapisz / Nie zapisuj / Anuluj”).
//
//   appDialog({ title, message, actions: [{ label, value, kind }], input?, content? })
//     → Promise<value | null>   (Esc, ✕, klik w tło = null)
//     kind: "primary" | "danger" | "secondary"; z `input` akcja primary zwraca wpisany tekst
//   appConfirm(message, { title, okLabel, cancelLabel, danger }) → Promise<boolean>
//   appPrompt(message, defaultValue, { title, okLabel, placeholder, validate }) → Promise<string|null>
//   appAskUnsaved({ title, message }) → Promise<"save" | "discard" | null>
//
// Fokus: z klawiatury (albo myszką na komputerze) siada w polu tekstowym albo na bezpiecznej
// akcji; na dotyku NIE przenosimy go — na iOS fokus w polu wysuwa klawiaturę, a fokus na
// przycisku zostawiał go „wciśniętym”. Tab krąży w obrębie okna, Esc = Anuluj, po zamknięciu
// fokus wraca tam, skąd okno otwarto. Klawisze nie przeciekają do tabeli pod spodem.

const appDialogStack = [];

function appDialogWantsFocus() {
  const root = document.documentElement;
  if (root.classList.contains("kbd-focus")) return true;
  return !(window.matchMedia && matchMedia("(pointer: coarse)").matches);
}

function appDialogFocusables(box) {
  return Array.from(box.querySelectorAll("button, input, select, textarea, a[href], [tabindex]:not([tabindex='-1'])"))
    .filter((el) => !el.disabled && el.offsetParent !== null);
}

function appDialog(opts = {}) {
  const {
    title = "",
    message = "",
    actions = [{ label: t("dlgOk"), value: true, kind: "primary" }, { label: t("dlgCancel"), value: null, kind: "secondary" }],
    input = null,
    content = null,
    className = "",
  } = opts;

  return new Promise((resolve) => {
    const returnTo = document.activeElement && document.activeElement !== document.body ? document.activeElement : null;
    const overlay = document.createElement("div");
    overlay.className = `modal app-dialog ${className}`.trim();
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    const uid = `appDlg${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;

    const box = document.createElement("div");
    box.className = "modal-content app-dialog-box";

    const head = document.createElement("div");
    head.className = "modal-head";
    const h = document.createElement("div");
    h.className = "modal-title";
    h.id = `${uid}T`;
    h.textContent = title || "";
    const x = document.createElement("button");
    x.type = "button";
    x.className = "modal-close";
    x.setAttribute("aria-label", t("close"));
    x.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';
    head.append(h, x);
    box.appendChild(head);
    if (title) overlay.setAttribute("aria-labelledby", h.id);

    const body = document.createElement("div");
    body.className = "app-dialog-body";
    if (message) {
      const p = document.createElement("p");
      p.className = "app-dialog-msg";
      p.id = `${uid}M`;
      p.textContent = message;
      body.appendChild(p);
      overlay.setAttribute("aria-describedby", p.id);
    }
    if (content instanceof Node) body.appendChild(content);

    let field = null;
    let err = null;
    if (input) {
      const label = document.createElement("label");
      label.className = "field app-dialog-field";
      const cap = document.createElement("span");
      cap.textContent = input.label || "";
      field = document.createElement("input");
      field.type = "text";
      field.autocomplete = "off";
      field.spellcheck = false;
      field.value = input.value != null ? String(input.value) : "";
      if (input.placeholder) field.placeholder = input.placeholder;
      if (input.maxLength) field.maxLength = input.maxLength;
      field.setAttribute("enterkeyhint", "done");
      if (input.label) label.appendChild(cap);
      else field.setAttribute("aria-label", title || message || "");
      label.appendChild(field);
      err = document.createElement("div");
      err.className = "app-dialog-error";
      err.setAttribute("role", "alert");
      err.hidden = true;
      body.append(label, err);
    }
    box.appendChild(body);

    const bar = document.createElement("div");
    bar.className = "modal-actions app-dialog-actions";
    const buttons = actions.map((a) => {
      const b = document.createElement("button");
      b.type = "button";
      const kind = a.kind || "secondary";
      b.className = kind === "primary" ? "btn" : kind === "danger" ? "btn ghost danger" : "btn ghost";
      b.dataset.kind = kind;
      b.dataset.value = String(a.value);
      b.textContent = a.label;
      b.addEventListener("click", () => choose(a));
      bar.appendChild(b);
      return { b, a };
    });
    box.appendChild(bar);
    overlay.appendChild(box);

    let done = false;
    function finish(value) {
      if (done) return;
      done = true;
      window.removeEventListener("keydown", onKey, true);
      const i = appDialogStack.indexOf(entry);
      if (i >= 0) appDialogStack.splice(i, 1);
      overlay.remove();
      if (returnTo && returnTo.isConnected && typeof returnTo.focus === "function") {
        try { returnTo.focus({ preventScroll: true }); } catch { /* element zniknął */ }
      }
      resolve(value);
    }
    function choose(a) {
      if (field && a.kind === "primary") {
        const text = field.value.trim();
        const problem = typeof input.validate === "function" ? input.validate(text) : "";
        if (problem) {
          err.textContent = problem;
          err.hidden = false;
          field.setAttribute("aria-invalid", "true");
          if (appDialogWantsFocus()) field.focus();
          return;
        }
        finish(text);
        return;
      }
      finish(a.value === undefined ? null : a.value);
    }
    if (field) field.addEventListener("input", () => { if (!err.hidden) { err.hidden = true; field.removeAttribute("aria-invalid"); } });

    x.addEventListener("click", () => finish(null));
    overlay.addEventListener("click", (e) => { if (e.target === overlay) finish(null); });

    // Faza capture na window: przed menu ⋯, globalnym Esc i nawigacją po tabeli.
    function onKey(e) {
      if (appDialogStack[appDialogStack.length - 1] !== entry) return;
      if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        finish(null);
        return;
      }
      if (e.key === "Tab") {
        const f = appDialogFocusables(box);
        if (!f.length) { e.preventDefault(); e.stopPropagation(); return; }
        const i = f.indexOf(document.activeElement);
        let next = null;
        if (i < 0) next = e.shiftKey ? f[f.length - 1] : f[0];
        else if (e.shiftKey && i === 0) next = f[f.length - 1];
        else if (!e.shiftKey && i === f.length - 1) next = f[0];
        if (next) { e.preventDefault(); next.focus(); }
        e.stopPropagation();
        return;
      }
      if (e.key === "Enter" && field && e.target === field && !e.isComposing) {
        e.preventDefault();
        e.stopPropagation();
        const primary = actions.find((a) => a.kind === "primary");
        if (primary) choose(primary);
        return;
      }
      // reszta klawiszy działa normalnie W OKNIE (pisanie, Spacja/Enter na przycisku),
      // ale nie trafia do skrótów aplikacji pod spodem
      if (!box.contains(e.target)) e.preventDefault();
      e.stopPropagation();
    }
    const entry = { overlay, finish };
    appDialogStack.push(entry);
    window.addEventListener("keydown", onKey, true);
    document.body.appendChild(overlay);

    if (appDialogWantsFocus()) {
      requestAnimationFrame(() => {
        if (done) return;
        if (field) { field.focus(); field.select(); return; }
        const safe = buttons.find(({ a }) => a.kind === "primary") || buttons.find(({ a }) => a.value === null || a.value === false) || buttons[0];
        // akcja niebezpieczna nigdy nie dostaje fokusu domyślnie (Enter nie może skasować)
        const pick = safe && safe.a.kind !== "danger" ? safe : buttons.find(({ a }) => a.value === null || a.value === false);
        if (pick) pick.b.focus();
      });
    } else {
      // fokus nie może zostać w polu pod spodem (Enter z klawiatury ekranowej poszedłby tam)
      if (document.activeElement && document.activeElement !== document.body && !box.contains(document.activeElement)) {
        try { document.activeElement.blur(); } catch { /* nic */ }
      }
    }
  });
}

async function appConfirm(message, opts = {}) {
  const v = await appDialog({
    title: opts.title || "",
    message,
    actions: [
      { label: opts.okLabel || t("dlgOk"), value: true, kind: opts.danger ? "danger" : "primary" },
      { label: opts.cancelLabel || t("dlgCancel"), value: null, kind: "secondary" },
    ],
  });
  return v === true;
}

function appPrompt(message, defaultValue = "", opts = {}) {
  return appDialog({
    title: opts.title || "",
    message,
    input: { value: defaultValue, placeholder: opts.placeholder || "", label: opts.label || "", maxLength: opts.maxLength || 0, validate: opts.validate },
    actions: [
      { label: opts.okLabel || t("dlgSaveName"), value: true, kind: "primary" },
      { label: t("dlgCancel"), value: null, kind: "secondary" },
    ],
  }).then((v) => (typeof v === "string" ? v : null));
}

// Niezapisana praca: zawsze „Zapisz / Nie zapisuj / Anuluj” (w tej kolejności, Zapisz = główna).
function appAskUnsaved({ title, message } = {}) {
  return appDialog({
    title: title || t("dlgUnsavedTitle"),
    message: message || t("dlgUnsavedMsg"),
    className: "app-dialog-unsaved",
    actions: [
      { label: t("dlgSave"), value: "save", kind: "primary" },
      { label: t("dlgDontSave"), value: "discard", kind: "danger" },
      { label: t("dlgCancel"), value: null, kind: "secondary" },
    ],
  });
}

// „Zapisz” z okna niezapisanej pracy: zapis jak przyciskiem Zapisz; dalej idziemy tylko,
// gdy zmiany naprawdę trafiły do pliku (anulowany wybór miejsca zapisu = zostajemy).
async function appSaveBeforeLeaving() {
  if (typeof saveWorkbook !== "function") return false;
  await saveWorkbook();
  return !hasUnsavedChanges;
}

// Wspólna bramka „co z niezapisanymi zmianami?” — true = można iść dalej.
async function appResolveUnsaved(opts = {}) {
  if (!hasUnsavedChanges) return true;
  const choice = await appAskUnsaved(opts);
  if (choice === "discard") return true;
  if (choice === "save") return appSaveBeforeLeaving();
  return false;
}
