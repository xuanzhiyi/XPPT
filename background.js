// XPPT background service worker
// Intercepts .pptx navigations and downloads → redirects to the viewer page.

const PPTX_RE = /\.pptx(\?[^#]*)?$/i;

// Track tabs we're already redirecting to avoid double-firing.
const redirecting = new Set();

function viewerUrl(pptxUrl) {
  return chrome.runtime.getURL('pptx_viewer.html') + '?url=' + encodeURIComponent(pptxUrl);
}

function maybeRedirect(tabId, url) {
  if (!PPTX_RE.test(url)) return;
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
  if (PPTX_RE.test(item.url) || PPTX_RE.test(item.filename || '')) {
    chrome.downloads.cancel(item.id, () => {
      chrome.downloads.erase({ id: item.id });
      chrome.tabs.create({ url: viewerUrl(item.url) });
    });
  }
});
