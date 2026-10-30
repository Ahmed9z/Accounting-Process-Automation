'use strict';

// NOTE: In MV3, service workers can be killed between events.
// Global variables like `let pendingDownloadPath` are lost when this happens.
// We use chrome.storage.local instead so the path survives restarts.

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {

  if (message.action === 'setNextDownloadPath') {
    chrome.storage.local.set({ pendingDownloadPath: message.filename }, () => {
      sendResponse({ ok: true });
    });
    return true;
  }

  if (message.action === 'download') {
    chrome.downloads.download({
      url: message.url,
      filename: message.filename,
      saveAs: false,
      conflictAction: 'uniquify'
    }, (downloadId) => {
      if (chrome.runtime.lastError) {
        sendResponse({ success: false, error: chrome.runtime.lastError.message });
      } else {
        sendResponse({ success: true, downloadId });
      }
    });
    return true;
  }
});

// Intercept downloads and save to the correct BL subfolder.
// Reads from storage (not global var) so it works after service worker restarts.
// Requires "Ask where to save each file" to be OFF in chrome://settings/downloads
chrome.downloads.onDeterminingFilename.addListener((downloadItem, suggest) => {
  chrome.storage.local.get(['pendingDownloadPath'], (result) => {
    if (result.pendingDownloadPath) {
      const path = result.pendingDownloadPath;
      chrome.storage.local.remove('pendingDownloadPath');
      suggest({ filename: path, conflictAction: 'uniquify' });
    } else {
      suggest({});
    }
  });
  return true; // REQUIRED: tells Chrome we call suggest() asynchronously
});
