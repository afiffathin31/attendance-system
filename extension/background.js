// Background Service Worker for Manifest V3
// Handles cross-origin requests to Google Apps Script bypassing page CSP
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'SYNC_ATTENDANCE') {
    fetch(message.targetUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(message.payload)
    })
    .then(async (res) => {
      const data = await res.json();
      sendResponse({ success: true, data });
    })
    .catch((err) => {
      sendResponse({ success: false, error: err.message });
    });

    return true; // Keep channel open for asynchronous sendResponse
  }

  if (message.action === 'START_BACKGROUND_TIMER') {
    const minutes = message.intervalMinutes || 5;
    chrome.alarms.create('AUTO_SYNC_ALARM', { periodInMinutes: minutes });
    sendResponse({ success: true });
    return false;
  }

  if (message.action === 'STOP_BACKGROUND_TIMER') {
    chrome.alarms.clear('AUTO_SYNC_ALARM');
    sendResponse({ success: true });
    return false;
  }
});

// When Chrome Alarm fires, notify Meet/Zoom tabs to trigger sync
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === 'AUTO_SYNC_ALARM') {
    chrome.tabs.query({}, (tabs) => {
      if (!tabs) return;
      tabs.forEach((tab) => {
        if (tab.url && (tab.url.includes('zoom.us') || tab.url.includes('meet.google.com'))) {
          chrome.tabs.sendMessage(tab.id, { action: 'TRIGGER_AUTO_SYNC' }).catch(() => {});
        }
      });
    });
  }
});
