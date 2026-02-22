// --- Email storage ---
const STORAGE_KEY = 'butter-mail-emails';
const EMBEDDINGS_KEY = 'butter-mail-embeddings';
const CATEGORIES_KEY = 'butter-mail-categories';
const PCA_KEY = 'butter-mail-pca';
const PCA_POINTS_KEY = 'butter-mail-pca-points';
const PROMPT_CLUSTERS_KEY = 'butter-mail-prompt-clusters';

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

function getEmbeddings() {
  try {
    const raw = localStorage.getItem(EMBEDDINGS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveEmbeddings(embeddings) {
  localStorage.setItem(EMBEDDINGS_KEY, JSON.stringify(embeddings));
}

function getCategories() {
  try {
    const raw = localStorage.getItem(CATEGORIES_KEY);
    return raw ? JSON.parse(raw) : { assignments: {}, meta: {} };
  } catch {
    return { assignments: {}, meta: {} };
  }
}

function saveCategories(cats) {
  localStorage.setItem(CATEGORIES_KEY, JSON.stringify(cats));
}

function getPcaModel() {
  try {
    const raw = localStorage.getItem(PCA_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function savePcaModel(model) {
  if (model) localStorage.setItem(PCA_KEY, JSON.stringify(model));
  else localStorage.removeItem(PCA_KEY);
}

function getPcaPoints() {
  try {
    const raw = localStorage.getItem(PCA_POINTS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function savePcaPoints(points) {
  if (points && Object.keys(points).length > 0) {
    localStorage.setItem(PCA_POINTS_KEY, JSON.stringify(points));
  } else {
    localStorage.removeItem(PCA_POINTS_KEY);
  }
}

function getPromptClusters() {
  try {
    const raw = localStorage.getItem(PROMPT_CLUSTERS_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function savePromptClusters(clusters) {
  localStorage.setItem(PROMPT_CLUSTERS_KEY, JSON.stringify(clusters));
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

  const msgId = (headers['message-id'] || '').trim();
  const inReplyTo = (headers['in-reply-to'] || '').trim();
  const refs = (headers.references || '').trim();
  return {
    id: 'eml-' + Date.now() + '-' + Math.random().toString(36).slice(2),
    from: headers.from || '',
    to: headers.to || '',
    subject: (headers.subject || '(no subject)').replace(/\s+/g, ' ').trim(),
    date: headers.date || '',
    body,
    fromEmail: extractEmail(headers.from),
    messageId: msgId,
    inReplyTo,
    references: refs
  };
}

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

let imapEmails = [];

function getAllEmails() {
  const imported = getEmails();
  return [...imapEmails, ...imported];
}

// --- State ---
let currentView = 'list';
let currentFilter = 'all';
let pcaPoints = getPcaPoints();
let searchQuery = '';
let searchResults = null;

function getFilteredEmails() {
  let emails = searchResults !== null ? searchResults : getAllEmails();
  if (currentFilter === 'all') return emails;
  if (currentFilter.startsWith('prompt-')) {
    const slug = currentFilter.slice(7);
    const clusters = getPromptClusters();
    const cluster = clusters[slug];
    if (!cluster || !cluster.emailIds) return [];
    const idSet = new Set(cluster.emailIds);
    return emails.filter((e) => idSet.has(e.id));
  }
  const cats = getCategories();
  return emails.filter((e) => cats.assignments[e.id] === currentFilter);
}

function getEmailsWithCategories(emails) {
  const cats = getCategories();
  return emails.map((e) => ({
    ...e,
    categoryId: cats.assignments[e.id] || null
  }));
}

window.getCategoryColor = function (catId) {
  const cats = getCategories();
  return cats.meta[catId] && cats.meta[catId].color ? cats.meta[catId].color : '#E5C94A';
};

// --- Compute embeddings ---
async function computeEmbeddings() {
  if (typeof window.electronAPI === 'undefined') {
    alert('Compute embeddings requires the Electron app. Run: npm start');
    return;
  }
  let emails = getAllEmails();
  if (emails.length === 0) {
    alert('No emails. Import or fetch from IMAP first.');
    return;
  }

  const imapWithoutBody = emails.filter((e) => e.id.startsWith('imap-') && !e.body);
  if (imapWithoutBody.length > 0) {
    const uids = imapWithoutBody.map((e) => e.uid).filter(Boolean);
    const res = await window.electronAPI.imap.fetchBodies(uids);
    if (res.ok && res.bodies) {
      emails = emails.map((e) => {
        if (e.id.startsWith('imap-') && res.bodies[e.uid]) {
          return { ...e, body: res.bodies[e.uid].body, bodyIsHtml: res.bodies[e.uid].bodyIsHtml };
        }
        return e;
      });
    }
  }

  const progressEl = document.getElementById('progress-text');
  progressEl.textContent = 'Loading model...';
  const btn = document.getElementById('compute-embeddings-btn');
  btn.disabled = true;

  if (window.electronAPI.embeddings.onProgress) {
    window.electronAPI.embeddings.onProgress((p) => {
      progressEl.textContent = `${p.current} / ${p.total} ${p.message || ''}`;
    });
  }

  try {
    const res = await window.electronAPI.embeddings.compute(emails);
    if (!res.ok) throw new Error(res.error || 'Failed');
    saveEmbeddings(res.embeddings);

    const emailIds = Object.keys(res.embeddings);
    if (emailIds.length > 0) {
      const pcaRes = await window.electronAPI.embeddings.pca(res.embeddings, emailIds);
      if (pcaRes.ok && pcaRes.points) {
        pcaPoints = pcaRes.points;
        savePcaModel(pcaRes.model);
        savePcaPoints(pcaPoints);
      }
      const clusterRes = await window.electronAPI.embeddings.cluster(res.embeddings, emailIds);
      if (clusterRes.ok) {
        saveCategories({ assignments: clusterRes.assignments, meta: clusterRes.meta });
      }
    }
    progressEl.textContent = 'Done.';
    updateSubtabBar();
    refreshCurrentView();
  } catch (err) {
    progressEl.textContent = '';
    alert('Error: ' + (err.message || String(err)));
  } finally {
    btn.disabled = false;
    setTimeout(() => { progressEl.textContent = ''; }, 2000);
  }
}

// --- Subtab bar ---
function updateSubtabBar() {
  const bar = document.getElementById('subtab-bar');
  if (!bar) return;
  const cats = getCategories();
  const ids = Object.keys(cats.meta || {}).filter((k) => k !== 'noise').sort();
  const promptClusters = getPromptClusters();
  const promptSlugs = Object.keys(promptClusters).sort();

  const activeTab = currentFilter;
  let html = '<button type="button" class="subtab-btn' + (activeTab === 'all' ? ' active' : '') + '" data-subtab="all">all</button>';
  ids.forEach((cid) => {
    const meta = cats.meta[cid];
    const name = meta && meta.name ? meta.name : cid;
    const isActive = activeTab === cid;
    html += '<button type="button" class="subtab-btn' + (isActive ? ' active' : '') + '" data-subtab="' + escapeHtml(cid) + '" data-category-id="' + escapeHtml(cid) + '">' + escapeHtml(name) + '</button>';
  });
  promptSlugs.forEach((slug) => {
    const cluster = promptClusters[slug];
    const label = cluster && cluster.label ? cluster.label : slug;
    const tabId = 'prompt-' + slug;
    const isActive = activeTab === tabId;
    html += '<button type="button" class="subtab-btn subtab-prompt' + (isActive ? ' active' : '') + '" data-subtab="' + escapeHtml(tabId) + '">' + escapeHtml(label) + '</button>';
  });
  bar.innerHTML = html;

  ids.forEach((cid) => {
    const btn = bar.querySelector('[data-subtab="' + cid + '"]');
    if (btn) btn.addEventListener('dblclick', () => renameCategory(cid, btn));
  });

  bar.querySelectorAll('.subtab-btn').forEach((b) => {
    b.addEventListener('click', () => {
      bar.querySelectorAll('.subtab-btn').forEach((x) => x.classList.remove('active'));
      b.classList.add('active');
      currentFilter = b.dataset.subtab;
      refreshCurrentView();
    });
  });
}

function renameCategory(catId, btnEl) {
  const name = prompt('Rename category:', btnEl.textContent);
  if (name === null || !name.trim()) return;
  const cats = getCategories();
  if (cats.meta[catId]) {
    cats.meta[catId].name = name.trim();
    saveCategories(cats);
    btnEl.textContent = name.trim();
  }
}

// --- View switching ---
let timelineHeadersLoaded = false;

async function ensureTimelineThreadHeaders() {
  if (typeof window.electronAPI === 'undefined' || timelineHeadersLoaded) return;
  const imapWithoutThread = imapEmails.filter((e) => !e.messageId && e.uid);
  if (imapWithoutThread.length === 0) {
    timelineHeadersLoaded = true;
    return;
  }
  const uids = imapWithoutThread.map((e) => e.uid).filter(Boolean);
  try {
    const res = await window.electronAPI.imap.fetchThreadHeaders(uids);
    if (!res.ok || !res.headers) return;
    Object.keys(res.headers).forEach((uid) => {
      const h = res.headers[uid];
      const email = imapEmails.find((e) => String(e.uid) === String(uid));
      if (email && h) {
        email.messageId = h.messageId || '';
        email.inReplyTo = h.inReplyTo || '';
        email.references = h.references || '';
      }
    });
    timelineHeadersLoaded = true;
  } catch (_) {}
}

function refreshCurrentView() {
  const emails = getFilteredEmails();
  if (currentView === 'list') {
    renderEmailList(emails);
  } else if (currentView === 'graph') {
    renderGraphView(emails);
  } else if (currentView === 'timeline') {
    if (window.TimelineView) {
      window.TimelineView.render('timeline-stacks', emails, openEmailDetail);
    }
    ensureTimelineThreadHeaders().then(() => {
      if (currentView === 'timeline' && window.TimelineView) {
        window.TimelineView.render('timeline-stacks', getFilteredEmails(), openEmailDetail);
      }
    });
  }
}

function renderGraphView(emails) {
  const container = document.getElementById('graph-container');
  if (!container || !window.GraphView) return;
  const embeddings = getEmbeddings();
  const filteredIds = emails.map((e) => e.id).filter((id) => embeddings[id] && pcaPoints[id]);
  const points = {};
  filteredIds.forEach((id) => {
    if (pcaPoints[id]) points[id] = pcaPoints[id];
  });
  const emailsById = {};
  emails.forEach((e) => { emailsById[e.id] = getEmailsWithCategories([e])[0]; });
  if (!window.GraphView.init) return;
  if (!window._graphInited) {
    window.GraphView.init('graph-container');
    window.GraphView.animate();
    window._graphInited = true;
  }
  window.GraphView.render(points, emailsById);
}

// --- Import ---
document.getElementById('email-import').addEventListener('change', (e) => {
  const files = e.target.files;
  if (!files || files.length === 0) return;

  const emails = getEmails();
  let added = 0;

  const readNext = (index) => {
    if (index >= files.length) {
      saveEmails(emails);
      refreshCurrentView();
      updateSubtabBar();
      e.target.value = '';
      if (added > 0) alert('Imported ' + added + ' email(s).');
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
          parsed.forEach((p) => { emails.push(p); added++; });
        } else {
          emails.push(parseEml(text));
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

// --- Fetch IMAP ---
async function fetchFromImap() {
  if (typeof window.electronAPI === 'undefined') return;
  const btn = document.getElementById('refresh-btn');
  btn.classList.add('loading');
  btn.disabled = true;
  try {
    const result = await window.electronAPI.imap.fetch(650);
    if (result.ok) {
      imapEmails = result.emails;
      timelineHeadersLoaded = false;
    } else if (!result.error.includes('not configured')) alert('IMAP error: ' + (result.error || 'Unknown'));
    refreshCurrentView();
    updateSubtabBar();
  } finally {
    btn.classList.remove('loading');
    btn.disabled = false;
  }
}

document.getElementById('refresh-btn').addEventListener('click', fetchFromImap);
document.getElementById('compute-embeddings-btn').addEventListener('click', computeEmbeddings);

async function createPromptCluster() {
  if (typeof window.electronAPI === 'undefined') {
    alert('Create cluster requires the Electron app. Run: npm start');
    return;
  }
  const input = document.getElementById('prompt-cluster-input');
  const prompt = (input && input.value || '').trim();
  if (!prompt) {
    alert('Enter a word or phrase (e.g. application) to create a cluster.');
    return;
  }
  const embeddings = getEmbeddings();
  const emailIds = Object.keys(embeddings);
  if (emailIds.length === 0) {
    alert('No embeddings. Run "compute embeddings" first.');
    return;
  }
  const btn = document.getElementById('prompt-cluster-btn');
  btn.disabled = true;
  const progressEl = document.getElementById('progress-text');
  progressEl.textContent = 'Creating cluster...';
  try {
    const res = await window.electronAPI.embeddings.promptCluster(prompt, embeddings, emailIds, 0.5);
    if (!res.ok) throw new Error(res.error || 'Failed');
    const slug = prompt.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'cluster';
    const clusters = getPromptClusters();
    clusters[slug] = { emailIds: res.emailIds, label: prompt, createdAt: new Date().toISOString() };
    savePromptClusters(clusters);
    if (input) input.value = '';
    updateSubtabBar();
    refreshCurrentView();
    progressEl.textContent = 'Created cluster "' + prompt + '" with ' + res.emailIds.length + ' email(s).';
  } catch (err) {
    progressEl.textContent = '';
    alert('Error: ' + (err.message || String(err)));
  } finally {
    btn.disabled = false;
    setTimeout(() => { progressEl.textContent = ''; }, 3000);
  }
}

document.getElementById('prompt-cluster-btn').addEventListener('click', createPromptCluster);
document.getElementById('prompt-cluster-input').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') createPromptCluster();
});

// --- Render email list ---
function renderEmailList(emails) {
  const listEl = document.getElementById('email-list');
  if (!listEl) return;

  const sorted = [...emails].sort((a, b) => {
    const da = new Date(a.date || 0);
    const db = new Date(b.date || 0);
    return db - da;
  });

  if (sorted.length === 0) {
    const hint = typeof window.electronAPI !== 'undefined'
      ? 'Click "refresh" to fetch from IMAP, or import .eml / .mbox files.'
      : 'Import .eml or .mbox files to get started.';
    listEl.innerHTML = '<p class="email-list-empty">No emails yet. ' + hint + '</p>';
    return;
  }

  const cats = getCategories();
  const emailsWithCat = getEmailsWithCategories(sorted);
  listEl.innerHTML = emailsWithCat
    .map((e) => {
      const fromStr = e.from || e.fromEmail || 'Unknown';
      const catColor = e.categoryId && cats.meta[e.categoryId] ? cats.meta[e.categoryId].color : '#E5C94A';
      const catTitle = e.categoryId && cats.meta[e.categoryId] ? (cats.meta[e.categoryId].name || '') : '';
      return '<div class="email-row" data-id="' + escapeHtml(e.id) + '">' +
        '<span class="email-row-category-square" style="background-color:' + escapeHtml(catColor) + '" title="' + escapeHtml(catTitle) + '" aria-hidden></span>' +
        '<span class="email-row-subject">' + escapeHtml(e.subject) + '</span>' +
        '<span class="email-row-from">' + escapeHtml(fromStr) + '</span>' +
        '<span class="email-row-date">' + escapeHtml(formatDate(e.date)) + '</span>' +
        '</div>';
    })
    .join('');

  listEl.querySelectorAll('.email-row').forEach((row) => {
    row.addEventListener('click', () => {
      const email = getAllEmails().find((x) => x.id === row.dataset.id);
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

const SANITIZE_OPTS = {
  ALLOWED_TAGS: ['p', 'br', 'a', 'strong', 'em', 'u', 'ul', 'ol', 'li', 'blockquote', 'hr', 'img', 'h1', 'h2', 'h3', 'div', 'span'],
  ALLOWED_ATTR: ['href', 'src', 'alt', 'title']
};

function linkify(text) {
  const urlRe = /(https?:\/\/[^\s<]+)/g;
  return escapeHtml(text).replace(urlRe, (m) => '<a href="' + escapeHtml(m) + '" target="_blank" rel="noopener">' + escapeHtml(m) + '</a>').replace(/\n/g, '<br>');
}

function renderEmailBody(body, bodyIsHtml) {
  if (!body) return '';
  if (typeof DOMPurify === 'undefined') return escapeHtml(body).replace(/\n/g, '<br>');
  if (bodyIsHtml) return DOMPurify.sanitize(body, SANITIZE_OPTS);
  return DOMPurify.sanitize(linkify(body), SANITIZE_OPTS);
}

async function openEmailDetail(email) {
  document.getElementById('email-detail-subject').textContent = email.subject;
  document.getElementById('email-detail-from').textContent = email.from || email.fromEmail || '';
  document.getElementById('email-detail-to').textContent = email.toDisplay || email.to || '';
  document.getElementById('email-detail-date').textContent = email.date ? new Date(email.date).toLocaleString() : '';
  const bodyEl = document.getElementById('email-detail-body');
  let body = email.body;
  let bodyIsHtml = email.bodyIsHtml || false;
  if (!body && email.id && email.id.startsWith('imap-') && email.uid && typeof window.electronAPI !== 'undefined') {
    bodyEl.textContent = 'Loading...';
    const result = await window.electronAPI.imap.fetchOne(email.uid);
    if (result.ok && result.email) {
      body = result.email.body;
      bodyIsHtml = result.email.bodyIsHtml || false;
      if (result.email.toDisplay) email.toDisplay = result.email.toDisplay;
    }
  }
  bodyEl.innerHTML = renderEmailBody(body || '', bodyIsHtml);
  document.getElementById('email-detail').classList.remove('hidden');
}

document.getElementById('email-detail-close').addEventListener('click', () => {
  document.getElementById('email-detail').classList.add('hidden');
});

document.getElementById('email-detail').addEventListener('click', (e) => {
  if (e.target.id === 'email-detail') e.target.classList.add('hidden');
});

window.onGraphPointClick = function (emailId) {
  const email = getAllEmails().find((e) => e.id === emailId);
  if (email) openEmailDetail(email);
};

// --- Search (hybrid) ---
let searchDebounce = null;
document.querySelector('.search-input').addEventListener('input', (e) => {
  searchQuery = e.target.value.trim();
  if (searchDebounce) clearTimeout(searchDebounce);
  searchDebounce = setTimeout(async () => {
    if (!searchQuery) {
      searchResults = null;
    } else if (typeof window.electronAPI !== 'undefined' && window.electronAPI.search && window.HybridSearch) {
      const emails = getAllEmails();
      const embeddings = getEmbeddings();
      searchResults = await window.HybridSearch.search(searchQuery, emails, embeddings, window.electronAPI);
    } else {
      searchResults = getAllEmails().filter((e) =>
        (e.subject && e.subject.toLowerCase().includes(searchQuery.toLowerCase())) ||
        (e.body && e.body.toLowerCase().includes(searchQuery.toLowerCase())) ||
        (e.from && e.from.toLowerCase().includes(searchQuery.toLowerCase()))
      );
    }
    refreshCurrentView();
    searchDebounce = null;
  }, 300);
});

// --- View tabs ---
document.querySelectorAll('.view-tab').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.view-tab').forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
    currentView = btn.dataset.view;
    document.getElementById('panel-list').classList.toggle('hidden', currentView !== 'list');
    document.getElementById('panel-graph').classList.toggle('hidden', currentView !== 'graph');
    document.getElementById('panel-timeline').classList.toggle('hidden', currentView !== 'timeline');
    refreshCurrentView();
  });
});

// --- Settings ---
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
updateSubtabBar();
refreshCurrentView();
if (typeof window.electronAPI !== 'undefined') {
  fetchFromImap();
}
