document.getElementById('open-viewer').addEventListener('click', () => {
  chrome.tabs.create({ url: chrome.runtime.getURL('pptx_viewer.html') });
  window.close();
});
