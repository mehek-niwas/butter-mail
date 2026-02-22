// --- Email storage ---
const STORAGE_KEY = 'butter-mail-emails';

function getEmails() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function saveEmails(emails) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(emails));
}

// --- Parse .eml content ---
function parseEml(text) {
  const lines = text.split(/\r?\n/);
  const headers = {};
  let bodyStart = -1;
  let i = 0;

  for (; i < lines.length; i++) {
    const line = lines[i];
    if (line === '') {
      bodyStart = i;
      break;
    }
    const match = line.match(/^([\w-]+):\s*(.*)$/i);
    if (match) {
      const key = match[1].toLowerCase();
      let val = match[2];
      while (i + 1 < lines.length && /^\s/.test(lines[i + 1])) {
        i++;
        val += lines[i].trim();
      }
      headers[key] = val;
    }
  }

  const body = bodyStart >= 0 ? lines.slice(bodyStart + 1).join('\n').trim() : '';
  function extractEmail(str) {
    if (!str) return '';
    const m = str.match(/<([^>]+)>/);
    return m ? m[1] : str.trim();
  }

  return {
    id: 'eml-' + Date.now() + '-' + Math.random().toString(36).slice(2),
    from: headers.from || '',
    to: headers.to || '',
    subject: (headers.subject || '(no subject)').replace(/\s+/g, ' ').trim(),
    date: headers.date || '',
    body,
    fromEmail: extractEmail(headers.from)
  };
}

// --- Parse mbox file (multiple emails in one file) ---
function parseMbox(text) {
  const emails = [];
  const blocks = text.split(/\r?\n(?=From )/);
  let idx = 0;
  for (const block of blocks) {
    const trimmed = block.trim();
    if (!trimmed) continue;
    const lines = trimmed.split(/\r?\n/);
    const msg = lines[0].match(/^From /) ? lines.slice(1).join('\n') : trimmed;
    if (!msg.trim()) continue;
    try {
      const parsed = parseEml(msg);
      parsed.id = 'mbox-' + Date.now() + '-' + idx++;
      emails.push(parsed);
    } catch (_) {}
  }
  return emails;
}

// --- IMAP emails (in-memory, from Electron) ---
let imapEmails = [];

function getAllEmails() {
  const imported = getEmails();
  return [...imapEmails, ...imported];
}

// --- Import files (.eml / .mbox) ---
document.getElementById('email-import').addEventListener('change', (e) => {
  const files = e.target.files;
  if (!files || files.length === 0) return;

  const emails = getEmails();
  let added = 0;

  const readNext = (index) => {
    if (index >= files.length) {
      saveEmails(emails);
      renderEmailList(getAllEmails());
      e.target.value = '';
      if (added > 0) alert(`Imported ${added} email(s).`);
      return;
    }

    const file = files[index];
    const ext = (file.name || '').toLowerCase();
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = reader.result;
        if (ext.endsWith('.mbox')) {
          const parsed = parseMbox(text);
          parsed.forEach((p) => {
            emails.push(p);
            added++;
          });
        } else {
          const parsed = parseEml(text);
          emails.push(parsed);
          added++;
        }
      } catch (err) {
        console.warn('Could not parse', file.name, err);
      }
      readNext(index + 1);
    };
    reader.readAsText(file);
  };

  readNext(0);
});

// --- Fetch from IMAP (Electron only) ---
async function fetchFromImap() {
  if (typeof window.electronAPI === 'undefined') return;
  const btn = document.getElementById('refresh-btn');
  btn.classList.add('loading');
  btn.disabled = true;
  try {
    const result = await window.electronAPI.imap.fetch(100);
    if (result.ok) {
      imapEmails = result.emails;
    } else if (!result.error.includes('not configured')) {
      alert('IMAP error: ' + (result.error || 'Unknown'));
    }
    renderEmailList(getAllEmails());
  } finally {
    btn.classList.remove('loading');
    btn.disabled = false;
  }
}

document.getElementById('refresh-btn').addEventListener('click', fetchFromImap);

// --- Render email list ---
function renderEmailList(emails, search = '') {
  const listEl = document.getElementById('email-list');
  const query = search.toLowerCase().trim();
  const filtered = query
    ? emails.filter((e) =>
        (e.subject && e.subject.toLowerCase().includes(query)) ||
        (e.from && e.from.toLowerCase().includes(query)) ||
        (e.fromEmail && e.fromEmail.toLowerCase().includes(query)) ||
        (e.body && e.body.toLowerCase().includes(query))
      )
    : emails;

  filtered.sort((a, b) => {
    const da = new Date(a.date || 0);
    const db = new Date(b.date || 0);
    return db - da;
  });

  if (filtered.length === 0) {
    const hint = typeof window.electronAPI !== 'undefined'
      ? 'Click "refresh" to fetch from IMAP, or import .eml / .mbox files.'
      : 'Import .eml or .mbox files to get started.';
    listEl.innerHTML = `<p class="email-list-empty">No emails yet. ${hint}</p>`;
    return;
  }

  listEl.innerHTML = filtered
    .map(
      (e) =>
        `<div class="email-row" data-id="${escapeHtml(e.id)}">
          <span class="email-row-subject">${escapeHtml(e.subject)}</span>
          <span class="email-row-from">${escapeHtml(e.from || e.fromEmail || 'Unknown')}</span>
          <span class="email-row-date">${escapeHtml(formatDate(e.date))}</span>
        </div>`
    )
    .join('');

  listEl.querySelectorAll('.email-row').forEach((row) => {
    row.addEventListener('click', () => {
      const email = emails.find((x) => x.id === row.dataset.id);
      if (email) openEmailDetail(email);
    });
  });
}

function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function formatDate(dateStr) {
  if (!dateStr) return '';
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString(undefined, {
      month: 'short',
      day: 'numeric',
      year: d.getFullYear() !== new Date().getFullYear() ? 'numeric' : undefined
    });
  } catch {
    return dateStr;
  }
}

// --- Email detail ---
async function openEmailDetail(email) {
  document.getElementById('email-detail-subject').textContent = email.subject;
  document.getElementById('email-detail-meta').textContent =
    `From: ${email.from || email.fromEmail}${email.date ? ' · ' + new Date(email.date).toLocaleString() : ''}`;
  const bodyEl = document.getElementById('email-detail-body');
  let body = email.body;
  if (!body && email.id?.startsWith('imap-') && email.uid && typeof window.electronAPI !== 'undefined') {
    bodyEl.textContent = 'Loading...';
    const result = await window.electronAPI.imap.fetchOne(email.uid);
    if (result.ok && result.email) body = result.email.body;
  }
  bodyEl.innerHTML = escapeHtml(body || '').replace(/\n/g, '<br>');
  document.getElementById('email-detail').classList.remove('hidden');
}

document.getElementById('email-detail-close').addEventListener('click', () => {
  document.getElementById('email-detail').classList.add('hidden');
});

document.getElementById('email-detail').addEventListener('click', (e) => {
  if (e.target.id === 'email-detail') e.target.classList.add('hidden');
});

// --- Search ---
document.querySelector('.search-input').addEventListener('input', (e) => {
  renderEmailList(getAllEmails(), e.target.value);
});

// --- Tab switching ---
document.querySelectorAll('.nav-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.nav-btn').forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
  });
});

// --- Subtab switching ---
document.querySelectorAll('.subtab-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.subtab-btn').forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
    const subtab = btn.dataset.subtab;
    document.getElementById('panel-all').classList.toggle('hidden', subtab !== 'all');
    document.getElementById('panel-categories').classList.toggle('hidden', subtab === 'all');
  });
});

// --- Settings (Electron only) ---
const settingsOverlay = document.getElementById('settings-overlay');
const settingsForm = document.getElementById('settings-form');

document.getElementById('settings-btn').addEventListener('click', async () => {
  if (typeof window.electronAPI === 'undefined') {
    alert('IMAP settings are available in the Electron app. Run: npm start');
    return;
  }
  const config = await window.electronAPI.imap.getConfig();
  if (config) {
    settingsForm.elements.host.value = config.host || '';
    settingsForm.elements.port.value = config.port || 993;
    settingsForm.elements.user.value = config.user || '';
    settingsForm.elements.pass.value = config.pass || '';
  }
  document.getElementById('settings-status').textContent = '';
  settingsOverlay.classList.remove('hidden');
});

document.getElementById('settings-close').addEventListener('click', () => {
  settingsOverlay.classList.add('hidden');
});

settingsOverlay.addEventListener('click', (e) => {
  if (e.target === settingsOverlay) settingsOverlay.classList.add('hidden');
});

document.getElementById('settings-test').addEventListener('click', async () => {
  if (typeof window.electronAPI === 'undefined') return;
  const status = document.getElementById('settings-status');
  const config = {
    host: settingsForm.elements.host.value.trim(),
    port: parseInt(settingsForm.elements.port.value, 10) || 993,
    user: settingsForm.elements.user.value.trim(),
    pass: settingsForm.elements.pass.value
  };
  if (!config.host || !config.user || !config.pass) {
    status.textContent = 'Fill in host, email, and password.';
    return;
  }
  status.textContent = 'Testing...';
  const result = await window.electronAPI.imap.test(config);
  status.textContent = result.ok ? 'Connection OK!' : 'Failed: ' + result.error;
});

settingsForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  if (typeof window.electronAPI === 'undefined') return;
  const config = {
    host: settingsForm.elements.host.value.trim(),
    port: parseInt(settingsForm.elements.port.value, 10) || 993,
    secure: true,
    user: settingsForm.elements.user.value.trim(),
    pass: settingsForm.elements.pass.value
  };
  await window.electronAPI.imap.saveConfig(config);
  document.getElementById('settings-status').textContent = 'Saved. Click Refresh to fetch emails.';
});

// --- Initial load ---
renderEmailList(getAllEmails());
if (typeof window.electronAPI !== 'undefined') {
  fetchFromImap();
}
