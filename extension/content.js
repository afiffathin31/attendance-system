// Content Script: Absensi Otomatis untuk Google Meet & Zoom
(function() {
  console.log('[AbsensiExtension] Content script loaded on:', window.location.href);

  let autoSyncInterval = null;
  let isMinimized = false;
  const sessionSeenParticipants = new Set();
  let lastMeetingLocation = window.location.pathname;

  function showToast(message, isError = false) {
    let toast = document.getElementById('absensi-toast-el');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'absensi-toast-el';
      toast.className = 'absensi-toast';
      document.body.appendChild(toast);
    }
    toast.className = `absensi-toast show ${isError ? 'error' : ''}`;
    toast.innerHTML = `${isError ? '⚠️' : '✅'} <span>${message}</span>`;

    setTimeout(() => {
      toast.className = 'absensi-toast';
    }, 4000);
  }

  function isJunkText(text) {
    if (!text || text.length < 2 || text.length > 40) return true;
    if (text.includes('\n')) return true;

    const lower = text.toLowerCase();
    const junkSubstrings = [
      'more_vert', 'visual_effects', 'frame_person', 'lock_person', 'call_end',
      'back_hand', 'closed_caption', 'mood', 'computer_arrow', 'videocam',
      'mic', 'push_pin', 'volume_up', 'opsi lainnya', 'latar belakang',
      'bingkai ulang', 'orang lain', 'tambahkan orang', 'telusuri orang',
      'dalam rapat ini', 'kontributor', 'penyelenggara', 'turn on', 'turn off',
      'mute', 'audio', 'video', 'semua mikrofon', 'bisukan', 'mute all',
      'angkat tangan', 'raise hand', 'tangan terangkat', 'sematkan', 'lepas pin',
      'keluarkan dari panggilan', 'remove from call', 'tampilkan semua orang',
      'show everyone', 'search for people', 'telusuri peserta'
    ];

    for (const j of junkSubstrings) {
      if (lower.includes(j)) return true;
    }

    // Ignore single snake_case icon words (e.g. "more_vert", "frame_person")
    if (text.includes('_') && !text.includes(' ')) return true;

    // Names rarely exceed 5 words
    if (text.split(/\s+/).length > 5) return true;

    return false;
  }

  function extractGoogleMeetParticipants() {
    // Reset akumulator jika pengguna berpindah ruang rapat
    if (lastMeetingLocation !== window.location.pathname) {
      lastMeetingLocation = window.location.pathname;
      sessionSeenParticipants.clear();
    }

    // 1. Ambil nama langsung dari video tiles (atribut data-self-name)
    const selfNameTiles = document.querySelectorAll('[data-self-name]');
    selfNameTiles.forEach(el => {
      const val = el.getAttribute('data-self-name');
      if (val) {
        let clean = val.replace(/\s*\((Anda|You|Host|Penyelenggara rapat|Penyelenggara|Co-host|Tamu|Guest)\)/gi, '').trim();
        if (!isJunkText(clean) && clean.length > 1) {
          sessionSeenParticipants.add(clean);
        }
      }
    });

    // 2. Ambil nama dari panel daftar peserta (People Drawer)
    const listItems = document.querySelectorAll('div[role="listitem"]');
    listItems.forEach(item => {
      // Prioritas A: span dengan class zWGUib, title, atau dir="auto"
      const candidateSpans = item.querySelectorAll('span.zWGUib, span[title], span[dir="auto"]');
      let found = false;
      for (const sp of candidateSpans) {
        if (sp.classList.contains('google-material-icons') || sp.getAttribute('aria-hidden') === 'true') continue;
        let text = (sp.getAttribute('title') || sp.innerText || '').split('\n')[0].trim();
        text = text.replace(/\s*\((Anda|You|Host|Penyelenggara rapat|Penyelenggara|Co-host|Tamu|Guest)\)/gi, '').trim();
        if (!isJunkText(text) && text.length > 1) {
          sessionSeenParticipants.add(text);
          found = true;
          break;
        }
      }

      // Prioritas B: Daun span teks jika Google Meet mengubah class selector
      if (!found) {
        const allSpans = item.querySelectorAll('span');
        for (const sp of allSpans) {
          if (sp.children.length > 0) continue;
          if (sp.classList.contains('google-material-icons') || sp.getAttribute('aria-hidden') === 'true') continue;
          let text = (sp.innerText || '').split('\n')[0].trim();
          text = text.replace(/\s*\((Anda|You|Host|Penyelenggara rapat|Penyelenggara|Co-host|Tamu|Guest)\)/gi, '').trim();
          if (!isJunkText(text) && text.length > 1) {
            sessionSeenParticipants.add(text);
            break;
          }
        }
      }
    });

    // 3. Fallback: label nama di pojok bawah video tiles
    const tileBadges = document.querySelectorAll('div[data-participant-id] span.zWGUib, div[data-participant-id] span[dir="auto"]');
    tileBadges.forEach(el => {
      if (el.classList.contains('google-material-icons') || el.getAttribute('aria-hidden') === 'true') return;
      let text = (el.innerText || '').split('\n')[0].trim();
      text = text.replace(/\s*\((Anda|You|Host|Penyelenggara rapat|Penyelenggara|Co-host|Tamu|Guest)\)/gi, '').trim();
      if (!isJunkText(text) && text.length > 1) {
        sessionSeenParticipants.add(text);
      }
    });

    return Array.from(sessionSeenParticipants);
  }

  // Deep crawler untuk menembus Shadow DOM & IFrames yang sering digunakan Zoom Web
  function getAllAccessibleRoots() {
    const roots = [];
    const visited = new Set();

    function collect(node) {
      if (!node || visited.has(node)) return;
      visited.add(node);
      roots.push(node);

      // Telusuri open shadowRoot jika ada
      try {
        const walker = document.createTreeWalker(node.body || node, NodeFilter.SHOW_ELEMENT);
        let el;
        while (el = walker.nextNode()) {
          if (el.shadowRoot && !visited.has(el.shadowRoot)) {
            collect(el.shadowRoot);
          }
        }
      } catch (e) {}

      // Telusuri iframes yang satu origin
      try {
        const iframes = (node.querySelectorAll ? node.querySelectorAll('iframe') : []);
        iframes.forEach(iframe => {
          try {
            if (iframe.contentDocument && !visited.has(iframe.contentDocument)) {
              collect(iframe.contentDocument);
            }
          } catch (err) {}
        });
      } catch (e) {}
    }

    collect(document);
    return roots;
  }

  function extractZoomParticipants() {
    // Reset akumulator jika berpindah ruang Zoom
    if (lastMeetingLocation !== window.location.pathname) {
      lastMeetingLocation = window.location.pathname;
      sessionSeenParticipants.clear();
    }
    const set = new Set();
    const roots = getAllAccessibleRoots();

    function cleanZoomName(raw) {
      if (!raw) return '';
      let text = String(raw).trim();

      // 1. Abaikan header/jumlah peserta (contoh: "Participants (1)", "Peserta (2)")
      if (/^(participants|peserta)\s*\(\d+\)/i.test(text)) return '';

      // 2. Bersihkan badge role Zoom seperti (Host, me), (Host), (Co-host), (me), (You), dll.
      text = text.replace(/\((Host|Co-host|Host,\s*me|You|Me|Guest|Penyelenggara rapat|Anda)[^)]*\)/gi, '').trim();

      // 3. Hapus tombol-tombol action Zoom yang menempel di ujung teks (Invite, Mute, Mute All, More, dll.)
      text = text.replace(/\b(invite|mute all|mute|unmute|more|rename|chat|audio|video|pin)\b/gi, '').trim();

      // 4. Hapus newline & rapikan spasi
      text = text.replace(/[\n\r\t]+/g, ' ').replace(/\s{2,}/g, ' ').trim();

      // 5. Nama orang harus memiliki minimal 3 huruf alfabet (mencegah inisial "J" atau angka)
      const alphaOnly = text.replace(/[^a-zA-Z]/g, '');
      if (alphaOnly.length < 3) return '';

      // 6. Hindari teks menu/UI Zoom
      const lower = text.toLowerCase();
      const junk = [
        'mute', 'unmute', 'more', 'invite', 'mute all', 'rename', 'chat', 
        'participants', 'audio', 'video', 'share screen', 'record', 'reactions',
        'leave', 'end', 'security', 'settings', 'apps', 'whiteboards', 'view',
        'speaker', 'gallery', 'fullscreen', 'host tools', 'meeting information',
        'close', 'switch to', 'preview'
      ];
      if (junk.some(j => lower === j || lower.startsWith(j + ' ') || lower.endsWith(' ' + j))) return '';

      // 7. Batas panjang nama wajar (3 - 40 karakter)
      if (text.length < 3 || text.length > 40) return '';

      return text;
    }

    // Strategi 1: Selector CSS Eksklusif untuk Nama Peserta (Hanya elemen teks nama, bukan kontainer)
    const selectors = [
      '.participants-item__display-name',
      '[class*="participants-item__display-name"]',
      'span.name-content',
      'span[id^="participant-name"]',
      '[class*="video-avatar__avatar-name"]',
      '[class*="video-tile__name"]',
      '[class*="avatar-name"]',
      '[class*="video-name"]',
      '[class*="participant-name"]',
      'span[class*="user-name"]'
    ];

    roots.forEach(root => {
      try {
        selectors.forEach(sel => {
          root.querySelectorAll(sel).forEach(el => {
            const text = cleanZoomName(el.innerText || el.getAttribute('title') || '');
            if (text) set.add(text);
          });
        });
      } catch (e) {}
    });

    // Strategi 2: Anchor Role Zoom - Jika ada elemen (Host, me), ambil HANYA elemen display name di item list
    roots.forEach(root => {
      try {
        const textElements = root.querySelectorAll('span, div');
        textElements.forEach(el => {
          const rawText = (el.innerText || '').trim();
          if (/\((Host|Host,\s*me|You|Me)\)/i.test(rawText) && rawText.length > 3 && rawText.length < 40) {
            let item = el.closest('li, [role="listitem"], [class*="participants-item"]');
            if (item) {
              const nameEl = item.querySelector('.participants-item__display-name, [class*="display-name"], span[class*="name"]');
              if (nameEl) {
                const cleaned = cleanZoomName(nameEl.innerText || nameEl.getAttribute('title') || '');
                if (cleaned) set.add(cleaned);
              }
            }
          }
        });
      } catch (e) {}
    });

    // Strategi 3: Pencocokan Pola Teks Murni "Nama - Dept" (Ketat: 1 tanda hubung, 3-25 char nama)
    roots.forEach(root => {
      try {
        const candidateEls = root.querySelectorAll('span');
        candidateEls.forEach(el => {
          if (el.children.length === 0) {
            const t = (el.innerText || '').trim();
            if (/^[A-Za-z\s.'`]{3,30}\s*-\s*[A-Za-z\s\/]{2,30}$/.test(t)) {
              const cleaned = cleanZoomName(t);
              if (cleaned) set.add(cleaned);
            }
          }
        });
      } catch (e) {}
    });

    set.forEach(name => sessionSeenParticipants.add(name));
    console.log('[Absensi Extension] Peserta Zoom Terdeteksi (Total Session):', Array.from(sessionSeenParticipants));
    return Array.from(sessionSeenParticipants);
  }

  let autoSyncActive = false;
  let autoSyncIntervalMinutes = 5;
  let nextSyncTimestamp = 0;
  let countdownTimerId = null;

  function findScrollContainer(isMeet) {
    if (isMeet) {
      // 1. Cek dari item peserta jika sudah ada yang ter-render di DOM
      const item = document.querySelector('div[role="listitem"]');
      if (item) {
        let parent = item.parentElement;
        while (parent && parent !== document.body) {
          const style = window.getComputedStyle(parent);
          if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && parent.scrollHeight > parent.clientHeight) {
            return parent;
          }
          parent = parent.parentElement;
        }
      }

      // 2. Cek selector kontainer umum Google Meet
      const candidates = [
        document.querySelector('div[role="list"]'),
        document.querySelector('div[role="tabpanel"]'),
        document.querySelector('aside[aria-label*="Orang" i]'),
        document.querySelector('aside[aria-label*="People" i]'),
        document.querySelector('aside')
      ];
      for (const el of candidates) {
        if (!el) continue;
        if (el.scrollHeight > el.clientHeight) return el;
        const scrollableChild = el.querySelector('[style*="overflow-y: auto"], [style*="overflow-y: scroll"], [style*="overflow: auto"], [style*="overflow: scroll"]');
        if (scrollableChild && scrollableChild.scrollHeight > scrollableChild.clientHeight) return scrollableChild;
      }
    } else {
      // Zoom Web Client
      const roots = getAllAccessibleRoots();
      for (const root of roots) {
        const item = root.querySelector('.participants-item__display-name, [class*="participants-item"], li[role="listitem"]');
        if (item) {
          let parent = item.parentElement;
          while (parent && parent !== document.body) {
            const style = window.getComputedStyle(parent);
            if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && parent.scrollHeight > parent.clientHeight) {
              return parent;
            }
            parent = parent.parentElement;
          }
        }
        const zoomList = root.querySelector('.participants-ul, .participants-wrapper, [class*="participants-list"], ul[role="list"]');
        if (zoomList && zoomList.scrollHeight > zoomList.clientHeight) return zoomList;
      }
    }
    return null;
  }

  async function autoScrollAndHarvest(isMeet, isZoom) {
    let closedPanelAutoOpened = false;
    let peopleBtnToRestore = null;

    // 1. Google Meet: Jika panel peserta belum terbuka, buka otomatis agar DOM listitem tersedia
    if (isMeet) {
      let container = findScrollContainer(true);
      if (!container) {
        const peopleBtn = document.querySelector(
          'button[aria-label*="Orang" i], button[aria-label*="People" i], button[aria-label*="semua orang" i], button[aria-label*="everyone" i], button[data-panel-id="1"]'
        );
        if (peopleBtn) {
          const isPressed = peopleBtn.getAttribute('aria-pressed') === 'true' || peopleBtn.classList.contains('qs41qe');
          if (!isPressed) {
            peopleBtn.click();
            closedPanelAutoOpened = true;
            peopleBtnToRestore = peopleBtn;
            await new Promise(r => setTimeout(r, 450));
          }
        }
      }
    }

    // 2. Zoom Web: Cek apakah tombol Participants tertutup
    if (isZoom) {
      let container = findScrollContainer(false);
      if (!container) {
        const roots = getAllAccessibleRoots();
        for (const root of roots) {
          const zoomPartBtn = root.querySelector('button[aria-label*="Participants" i], button[aria-label*="Peserta" i], .footer-button__participants');
          if (zoomPartBtn && !zoomPartBtn.classList.contains('active') && zoomPartBtn.getAttribute('aria-expanded') !== 'true') {
            zoomPartBtn.click();
            closedPanelAutoOpened = true;
            peopleBtnToRestore = zoomPartBtn;
            await new Promise(r => setTimeout(r, 450));
            break;
          }
        }
      }
    }

    // 3. Panen peserta yang tampak saat ini
    if (isMeet) extractGoogleMeetParticipants();
    else if (isZoom) extractZoomParticipants();

    // 4. Jalankan virtual scrolling crawl dari atas sampai paling bawah
    const container = findScrollContainer(isMeet);
    if (container && container.scrollHeight > container.clientHeight) {
      const originalScrollTop = container.scrollTop;
      const clientHeight = container.clientHeight;
      const step = Math.max(Math.floor(clientHeight * 0.75), 150);

      let currentTop = 0;
      let maxScroll = container.scrollHeight - container.clientHeight;

      while (currentTop < maxScroll) {
        currentTop = Math.min(currentTop + step, maxScroll);
        container.scrollTop = currentTop;
        container.dispatchEvent(new Event('scroll', { bubbles: true }));
        await new Promise(r => setTimeout(r, 80));

        if (isMeet) extractGoogleMeetParticipants();
        else if (isZoom) extractZoomParticipants();

        maxScroll = container.scrollHeight - container.clientHeight;
      }

      // Pastikan mencapai titik paling bawah untuk menangkap user terakhir
      container.scrollTop = container.scrollHeight;
      container.dispatchEvent(new Event('scroll', { bubbles: true }));
      await new Promise(r => setTimeout(r, 100));

      if (isMeet) extractGoogleMeetParticipants();
      else if (isZoom) extractZoomParticipants();

      // Kembalikan posisi scroll pengguna ke titik semula secara transparan
      container.scrollTop = originalScrollTop;
      container.dispatchEvent(new Event('scroll', { bubbles: true }));
    }

    // 5. Tutup kembali panel jika tadi dibuka otomatis oleh skrip
    if (closedPanelAutoOpened && peopleBtnToRestore) {
      await new Promise(r => setTimeout(r, 150));
      const closeBtn = document.querySelector('button[aria-label*="Tutup" i], button[aria-label*="Close" i]');
      if (closeBtn && closeBtn.offsetParent !== null) {
        closeBtn.click();
      } else {
        peopleBtnToRestore.click();
      }
    }
  }

  async function performAttendanceSync(isAuto = false) {
    const isMeet = window.location.hostname.includes('meet.google.com');
    const isZoom = window.location.hostname.includes('zoom.us');

    // Auto-scroll virtual DOM untuk mengumpulkan semua peserta dari atas sampai paling bawah
    await autoScrollAndHarvest(isMeet, isZoom);

    const participants = isMeet ? extractGoogleMeetParticipants() : (isZoom ? extractZoomParticipants() : []);

    const countBadge = document.getElementById('absensi-count-badge');
    if (countBadge) countBadge.textContent = `${participants.length} jamaah`;

    if (participants.length === 0) {
      if (!isAuto) {
        showToast('Belum ada peserta terdeteksi. Pastikan tab "Orang/Peserta" terbuka.', true);
      }
      const syncStatusEl = document.getElementById('absensi-last-sync-status');
      if (syncStatusEl && isAuto) {
        syncStatusEl.textContent = '⚠️ Panel peserta tertutup (0 terdeteksi)';
        syncStatusEl.style.display = 'block';
      }
      return;
    }

    let sessionName = document.title ? document.title.replace('- Google Meet', '').trim() : 'Kajian Rutin';
    if (sessionName.toLowerCase().includes('zoom meeting') || sessionName.toLowerCase().includes('personal meeting')) {
      sessionName = 'Kajian Rutin';
    }
    const now = new Date();
    const scheduledStartTime = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

    // Read configured Cloud URL (Google Apps Script)
    const DEFAULT_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbybAeK8aGD8zGBonYH53SqxHkZqf95nLNBNSMJxfVhBGLfCwUdmn-VG_4-njdMhuNRCfg/exec';
    let targetUrl = DEFAULT_WEB_APP_URL;

    if (chrome && chrome.storage && chrome.storage.local) {
      try {
        const stored = await new Promise(r => chrome.storage.local.get(['webAppUrl'], r));
        if (stored && stored.webAppUrl && stored.webAppUrl.startsWith('http')) {
          targetUrl = stored.webAppUrl.trim();
        }
      } catch (e) {}
    }

    try {
      const payload = {
        participants,
        sessionName,
        scheduledStartTime,
        toleranceMinutes: 15,
        platform: isMeet ? 'GOOGLE_MEET' : 'ZOOM'
      };

      let data = null;

      // Prioritas 1: Kirim via Background Service Worker (bebas batasan CSP / CORS Zoom)
      if (chrome && chrome.runtime && chrome.runtime.sendMessage) {
        try {
          const bgResponse = await new Promise((resolve) => {
            chrome.runtime.sendMessage({
              action: 'SYNC_ATTENDANCE',
              targetUrl,
              payload
            }, (res) => {
              if (chrome.runtime.lastError || !res) {
                resolve(null);
              } else {
                resolve(res);
              }
            });
          });

          if (bgResponse && bgResponse.success && bgResponse.data) {
            data = bgResponse.data;
          }
        } catch (e) {}
      }

      // Prioritas 2: Fallback direct fetch jika background worker tidak merespons
      if (!data) {
        const res = await fetch(targetUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(payload)
        });
        data = await res.json();
      }

      if (data && data.success) {
        const kajianDetail = data.pemateri ? `\n📖 ${data.pemateri} - ${data.tema}` : '';
        const timeStr = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        const syncStatusEl = document.getElementById('absensi-last-sync-status');
        
        let statusText = '';
        if (data.count === 0 && participants.length > 0) {
          statusText = `🕒 Terakhir: ${timeStr} (${participants.length} jamaah verified)`;
          if (!isAuto) {
            showToast(`ℹ️ ${participants.length} jamaah sudah pernah tercatat hadir untuk kajian hari ini!${kajianDetail}`);
          }
        } else {
          statusText = `🕒 Terakhir: ${timeStr} (+${data.count} jamaah baru)`;
          showToast(`✅ Tercatat: ${data.count !== undefined ? data.count : participants.length} jamaah tersinkron!${kajianDetail}`);
        }

        if (syncStatusEl) {
          syncStatusEl.textContent = statusText;
          syncStatusEl.style.display = 'block';
        }

        if (chrome && chrome.storage && chrome.storage.local) {
          chrome.storage.local.set({ lastSyncText: statusText });
        }
      } else {
        showToast((data && data.error) || 'Gagal menyimpan absensi', true);
      }
    } catch (err) {
      showToast('Gagal mengirim ke Google Sheets. Pastikan akses Web App diatur ke "Anyone" (Siapa saja).', true);
    }
  }

  function updateCountdownUI() {
    const label = document.getElementById('absensi-countdown-label');
    if (!label) return;

    if (!autoSyncActive) {
      label.textContent = 'Nonaktif';
      label.classList.remove('active');
      return;
    }

    const remainingMs = Math.max(0, nextSyncTimestamp - Date.now());
    const totalSec = Math.floor(remainingMs / 1000);
    const m = String(Math.floor(totalSec / 60)).padStart(2, '0');
    const s = String(totalSec % 60).padStart(2, '0');

    label.textContent = `⏳ Sync dlm ${m}:${s}`;
    label.classList.add('active');

    if (remainingMs <= 0) {
      nextSyncTimestamp = Date.now() + autoSyncIntervalMinutes * 60 * 1000;
      performAttendanceSync(true);
    }
  }

  function startAutoSync(intervalMinutes) {
    autoSyncActive = true;
    if (intervalMinutes) {
      autoSyncIntervalMinutes = intervalMinutes;
    }
    nextSyncTimestamp = Date.now() + autoSyncIntervalMinutes * 60 * 1000;

    if (chrome && chrome.storage && chrome.storage.local) {
      chrome.storage.local.set({
        autoSyncActive: true,
        autoSyncIntervalMinutes: autoSyncIntervalMinutes
      });
    }

    if (chrome && chrome.runtime && chrome.runtime.sendMessage) {
      chrome.runtime.sendMessage({
        action: 'START_BACKGROUND_TIMER',
        intervalMinutes: autoSyncIntervalMinutes
      }).catch(() => {});
    }

    if (countdownTimerId) clearInterval(countdownTimerId);
    countdownTimerId = setInterval(updateCountdownUI, 1000);
    updateCountdownUI();

    // Langsung lakukan sync pertama saat diaktifkan agar pengguna tidak menunggu interval pertama
    performAttendanceSync(true);
  }

  function stopAutoSync() {
    autoSyncActive = false;
    if (countdownTimerId) {
      clearInterval(countdownTimerId);
      countdownTimerId = null;
    }

    if (chrome && chrome.storage && chrome.storage.local) {
      chrome.storage.local.set({ autoSyncActive: false });
    }

    if (chrome && chrome.runtime && chrome.runtime.sendMessage) {
      chrome.runtime.sendMessage({ action: 'STOP_BACKGROUND_TIMER' }).catch(() => {});
    }

    updateCountdownUI();
  }

  // Listener untuk pesan TRIGGER_AUTO_SYNC dari service worker background
  if (chrome && chrome.runtime && chrome.runtime.onMessage) {
    chrome.runtime.onMessage.addListener((message) => {
      if (message.action === 'TRIGGER_AUTO_SYNC' && autoSyncActive) {
        console.log('[Absensi Extension] Menerima sinyal alarm background AUTO_SYNC');
        nextSyncTimestamp = Date.now() + autoSyncIntervalMinutes * 60 * 1000;
        performAttendanceSync(true);
      }
    });
  }

  function injectWidget() {
    if (window !== window.top) return;
    if (document.getElementById('absensi-floating-widget')) return;

    const widget = document.createElement('div');
    widget.id = 'absensi-floating-widget';
    widget.innerHTML = `
      <div class="absensi-header">
        <div class="absensi-title">
          <span>📋 Absensi Rapat</span>
          <span id="absensi-count-badge" class="absensi-badge">Siap</span>
        </div>
        <button id="absensi-toggle-min-btn" class="absensi-minimize-btn" title="Kecilkan / Besarkan">−</button>
      </div>
      <div class="absensi-body">
        <button id="absensi-btn-trigger" class="absensi-btn-rekap">
          <span>⚡ Rekap Kehadiran Sekarang</span>
        </button>
        <div class="absensi-auto-sync-row">
          <div class="absensi-auto-sync-info">
            <span>Auto-Sync Berkala</span>
            <span id="absensi-countdown-label" class="absensi-countdown-label">Nonaktif</span>
            <select id="absensi-interval-select" class="absensi-interval-select" title="Pilih interval waktu auto-sync">
              <option value="1">1 Mnt (Tes Cepat)</option>
              <option value="2">2 Menit</option>
              <option value="5" selected>5 Menit (Default)</option>
              <option value="10">10 Menit</option>
            </select>
          </div>
          <label class="absensi-switch">
            <input type="checkbox" id="absensi-auto-toggle">
            <span class="absensi-slider"></span>
          </label>
        </div>
        <div id="absensi-last-sync-status" class="absensi-last-sync-status"></div>
      </div>
    `;

    document.body.appendChild(widget);

    // Event Listeners
    const rekapBtn = document.getElementById('absensi-btn-trigger');
    rekapBtn.addEventListener('click', () => {
      rekapBtn.innerHTML = '<span>⏳ Merekap data...</span>';
      performAttendanceSync(false).finally(() => {
        rekapBtn.innerHTML = '<span>⚡ Rekap Kehadiran Sekarang</span>';
      });
    });

    const minBtn = document.getElementById('absensi-toggle-min-btn');
    minBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      isMinimized = !isMinimized;
      widget.classList.toggle('minimized', isMinimized);
      minBtn.textContent = isMinimized ? '+' : '−';
    });

    widget.addEventListener('click', () => {
      if (isMinimized) {
        isMinimized = false;
        widget.classList.remove('minimized');
        minBtn.textContent = '−';
      }
    });

    const autoToggle = document.getElementById('absensi-auto-toggle');
    const intervalSelect = document.getElementById('absensi-interval-select');
    const lastSyncStatus = document.getElementById('absensi-last-sync-status');

    // Muat konfigurasi tersimpan
    if (chrome && chrome.storage && chrome.storage.local) {
      chrome.storage.local.get(['autoSyncActive', 'autoSyncIntervalMinutes', 'lastSyncText'], (res) => {
        if (res && res.lastSyncText && lastSyncStatus) {
          lastSyncStatus.textContent = res.lastSyncText;
          lastSyncStatus.style.display = 'block';
        }
        if (res && res.autoSyncIntervalMinutes && intervalSelect) {
          intervalSelect.value = String(res.autoSyncIntervalMinutes);
          autoSyncIntervalMinutes = parseInt(res.autoSyncIntervalMinutes, 10);
        }
        if (res && res.autoSyncActive && autoToggle) {
          autoToggle.checked = true;
          startAutoSync(autoSyncIntervalMinutes);
        }
      });
    }

    autoToggle.addEventListener('change', () => {
      if (autoToggle.checked) {
        const mins = parseInt(intervalSelect.value, 10) || 5;
        startAutoSync(mins);
        showToast(`Auto-sync aktif: mengecek kehadiran setiap ${mins} menit.`);
      } else {
        stopAutoSync();
        showToast('Auto-sync dinonaktifkan.');
      }
    });

    intervalSelect.addEventListener('change', () => {
      const mins = parseInt(intervalSelect.value, 10) || 5;
      autoSyncIntervalMinutes = mins;
      if (autoToggle.checked) {
        startAutoSync(mins);
        showToast(`Interval auto-sync diubah ke ${mins} menit.`);
      } else {
        if (chrome && chrome.storage && chrome.storage.local) {
          chrome.storage.local.set({ autoSyncIntervalMinutes: mins });
        }
      }
    });
  }

  // Check periodically to inject when meeting UI loads
  setInterval(() => {
    // Only inject if on Google Meet or Zoom
    if (window.location.hostname.includes('meet.google.com') || window.location.hostname.includes('zoom.us')) {
      injectWidget();
    }
  }, 2000);

  // Live counter & background harvester berkala (tiap 4 detik)
  setInterval(() => {
    const isMeet = window.location.hostname.includes('meet.google.com');
    const isZoom = window.location.hostname.includes('zoom.us');
    if (!isMeet && !isZoom) return;

    if (isMeet) extractGoogleMeetParticipants();
    else if (isZoom) extractZoomParticipants();

    const countBadge = document.getElementById('absensi-count-badge');
    if (countBadge && sessionSeenParticipants.size > 0) {
      countBadge.textContent = `${sessionSeenParticipants.size} jamaah`;
    }
  }, 4000);

})();

