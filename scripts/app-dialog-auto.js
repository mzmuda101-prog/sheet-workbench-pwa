// app-dialog-auto.js — odpowiednik `page.on("dialog", d => d.accept())` dla WŁASNYCH okien
// aplikacji (app/app-dialog.js, od 2026-11-07 zamiast confirm()/prompt()).
//   const { autoAnswerAppDialogs } = require("./app-dialog-auto");
//   await autoAnswerAppDialogs(page);          // przed page.goto
// W stronie: window.__appDialogMode = "accept" (domyślnie) | "dismiss" | "off";
//            window.__appDialogText = "nazwa"  — tekst wpisywany w okno z polem;
//            window.__appDialogLog            — [{ title, msg, buttons }] każdego okna.
// „accept” = to, co dawniej OK w confirm(): przy niezapisanych zmianach „Nie zapisuj”
// (porzuć i idź dalej), w pozostałych oknach przycisk główny.
function install() {
  const answer = (overlay) => {
    if (!overlay.isConnected) return;
    const mode = window.__appDialogMode || "accept";
    const btns = Array.from(overlay.querySelectorAll(".app-dialog-actions button"));
    (window.__appDialogLog = window.__appDialogLog || []).push({
      title: overlay.querySelector(".modal-title")?.textContent || "",
      msg: overlay.querySelector(".app-dialog-msg")?.textContent || "",
      buttons: btns.map((b) => b.textContent),
    });
    if (mode === "off") return;
    if (mode === "dismiss") { overlay.querySelector(".modal-close")?.click(); return; }
    const input = overlay.querySelector(".app-dialog-field input");
    if (input && typeof window.__appDialogText === "string") input.value = window.__appDialogText;
    const pick = btns.find((b) => b.dataset.value === "discard") || btns.find((b) => b.dataset.kind === "primary") || btns[0];
    pick?.click();
  };
  const start = () => new MutationObserver((recs) => {
    recs.forEach((r) => r.addedNodes.forEach((n) => {
      if (n.nodeType === 1 && n.classList.contains("app-dialog") && !n.classList.contains("sc-dialog")) setTimeout(() => answer(n), 0);
    }));
  }).observe(document.body, { childList: true });
  if (document.body) start(); else document.addEventListener("DOMContentLoaded", start);
}

async function autoAnswerAppDialogs(page) {
  await page.addInitScript(install);
}

module.exports = { autoAnswerAppDialogs, install };
