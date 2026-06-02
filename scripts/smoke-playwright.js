const { chromium } = require("playwright");

const APP_URL = process.env.APP_URL || "http://127.0.0.1:4175/";

function rectToPlain(rect) {
  if (!rect) return null;
  return {
    x: Math.round(rect.x),
    y: Math.round(rect.y),
    width: Math.round(rect.width),
    height: Math.round(rect.height),
    bottom: Math.round(rect.bottom ?? rect.y + rect.height),
  };
}

async function run() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: { width: 390, height: 844 },
    isMobile: true,
    hasTouch: true,
    serviceWorkers: "block",
  });
  await context.addInitScript(() => localStorage.setItem("introPlayed", "true"));

  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  page.on("console", (message) => {
    if (message.type() === "error") errors.push(message.text());
  });

  await page.goto(APP_URL, { waitUntil: "load" });
  const title = await page.title();
  await page.evaluate(() => {
    document.getElementById("heroSplash")?.remove();
  });
  await page.waitForTimeout(800);

  await page.evaluate(() => {
    if (typeof setSidebarOpen === "function") setSidebarOpen(false);
    else document.documentElement.classList.remove("sidebar-open");
  });
  await page.waitForTimeout(350);

  const closedAfterManualClose = await page
    .locator("html")
    .evaluate((el) => el.classList.contains("sidebar-open"));

  await page.locator("#panelHandle").tap({ force: true });
  await page.waitForTimeout(350);

  const openAfterClick = await page
    .locator("html")
    .evaluate((el) => el.classList.contains("sidebar-open"));
  const openRects = await page.evaluate(() => {
    const handle = document.getElementById("panelHandle");
    const sidebar = document.querySelector(".sidebar");
    const handleRect = handle.getBoundingClientRect();
    const sidebarRect = sidebar.getBoundingClientRect();
    return {
      innerWidth,
      innerHeight,
      handleRect: {
        x: handleRect.x,
        y: handleRect.y,
        width: handleRect.width,
        height: handleRect.height,
        bottom: handleRect.bottom,
      },
      sidebarRect: {
        x: sidebarRect.x,
        y: sidebarRect.y,
        width: sidebarRect.width,
        height: sidebarRect.height,
        bottom: sidebarRect.bottom,
      },
      handleLabel: getComputedStyle(handle, "::after").content,
    };
  });

  await page.locator("#panelHandle").tap({ force: true });
  await page.waitForTimeout(350);
  const closedAfterSecondClick = await page
    .locator("html")
    .evaluate((el) => el.classList.contains("sidebar-open"));

  await context.close();

  const serviceWorkerContext = await browser.newContext();
  const serviceWorkerPage = await serviceWorkerContext.newPage();
  await serviceWorkerPage.goto(APP_URL, { waitUntil: "load" });
  const serviceWorker = await serviceWorkerPage.evaluate(async () => {
    if (!("serviceWorker" in navigator)) return { supported: false };
    const registration = await navigator.serviceWorker.register("sw.js?v=20260602-05");
    await navigator.serviceWorker.ready;
    return {
      supported: true,
      scope: registration.scope,
      active: !!registration.active || !!registration.installing || !!registration.waiting,
    };
  });
  await serviceWorkerContext.close();

  await browser.close();

  const result = {
    title,
    closedAfterManualClose,
    openAfterClick,
    closedAfterSecondClick,
    handleRect: rectToPlain(openRects.handleRect),
    sidebarRect: rectToPlain(openRects.sidebarRect),
    viewport: {
      width: openRects.innerWidth,
      height: openRects.innerHeight,
    },
    handleLabel: openRects.handleLabel,
    serviceWorker,
    errors,
  };

  const failures = [];
  if (closedAfterManualClose) failures.push("sidebar should close before the handle-open check");
  if (!openAfterClick) failures.push("mobile handle should open the sidebar");
  if (closedAfterSecondClick) failures.push("mobile handle should close the sidebar on second click");
  if (!serviceWorker.supported || !serviceWorker.active) failures.push("service worker should be active");
  if (errors.length) failures.push(`console/page errors: ${errors.join(" | ")}`);

  console.log(JSON.stringify(result, null, 2));
  if (failures.length) {
    throw new Error(failures.join("; "));
  }
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
