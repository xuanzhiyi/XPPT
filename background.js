// XPPT background service worker
// Intercepts .pptx navigations and downloads → redirects to the viewer page.

// Match .pptx only in the URL path (before any ? or #), not in query-string values.
const PPTX_PATH_RE = /\.pptx([?#]|$)/i;

// Track tabs we're already redirecting to avoid double-firing.
const redirecting = new Set();

// The base URL of our viewer page — never redirect this.
const VIEWER_BASE = chrome.runtime.getURL('pptx_viewer.html');

function viewerUrl(pptxUrl) {
  return VIEWER_BASE + '?url=' + encodeURIComponent(pptxUrl);
}

function maybeRedirect(tabId, url) {
  if (!PPTX_PATH_RE.test(url)) return;
  // Never redirect our own viewer page (guards against double-wrapping).
  if (url.startsWith(VIEWER_BASE)) return;
  // Skip internal Chrome/extension URLs.
  if (url.startsWith('chrome://') || url.startsWith('chrome-extension://')) return;
  if (redirecting.has(tabId)) return;
  redirecting.add(tabId);
  chrome.tabs.update(tabId, { url: viewerUrl(url) }, () => {
    // Remove guard after navigation settles
    setTimeout(() => redirecting.delete(tabId), 3000);
  });
}

// 1. webNavigation fires for http/https .pptx links in pages
chrome.webNavigation.onBeforeNavigate.addListener((details) => {
  if (details.frameId !== 0) return; // main frame only
  maybeRedirect(details.tabId, details.url);
});

// 2. tabs.onUpdated catches file:// URLs opened from Windows Explorer
//    (Chrome may not fire webNavigation for file:// before the tab commits)
chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (changeInfo.url) maybeRedirect(tabId, changeInfo.url);
});

// 3. downloads.onCreated catches download-triggered .pptx files
chrome.downloads.onCreated.addListener((item) => {
  if (PPTX_PATH_RE.test(item.url) || PPTX_PATH_RE.test(item.filename || '')) {
    chrome.downloads.cancel(item.id, () => {
      chrome.downloads.erase({ id: item.id });
      chrome.tabs.create({ url: viewerUrl(item.url) });
    });
  }
});
