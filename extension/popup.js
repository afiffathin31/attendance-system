document.addEventListener('DOMContentLoaded', () => {
  const urlInput = document.getElementById('cloudUrl');
  const saveBtn = document.getElementById('saveUrlBtn');
  const toast = document.getElementById('saveToast');

  const DEFAULT_URL = 'https://script.google.com/macros/s/AKfycbybAeK8aGD8zGBonYH53SqxHkZqf95nLNBNSMJxfVhBGLfCwUdmn-VG_4-njdMhuNRCfg/exec';

  // Load saved URL from storage
  if (chrome && chrome.storage && chrome.storage.local) {
    chrome.storage.local.get(['webAppUrl'], (res) => {
      urlInput.value = (res && res.webAppUrl) ? res.webAppUrl : DEFAULT_URL;
    });
  } else {
    urlInput.value = DEFAULT_URL;
  }

  saveBtn.addEventListener('click', () => {
    const val = urlInput.value.trim();
    if (!val) {
      alert('Silakan masukkan URL Web App Google Apps Script Anda.');
      return;
    }
    if (chrome && chrome.storage && chrome.storage.local) {
      chrome.storage.local.set({ webAppUrl: val }, () => {
        toast.style.display = 'block';
        setTimeout(() => { toast.style.display = 'none'; }, 3000);
      });
    }
  });
});
