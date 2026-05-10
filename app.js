const STORAGE_KEY = 'butter-mail-emails';
const EMBEDDINGS_KEY = 'butter-mail-embeddings';
const PCA_KEY = 'butter-mail-pca';
const PCA_POINTS_KEY = 'butter-mail-pca-points';
const PROMPT_CLUSTERS_KEY = 'butter-mail-prompt-clusters';
const BOOKMARK_OVERRIDES_KEY = 'butter-mail-bookmark-overrides';
const THEME_KEY = 'butter-mail-theme';
const IMAP_CACHE_DB_NAME = 'butter-mail-imap-cache';
const IMAP_CACHE_STORE = 'cache';
const HOME_TAB_ID = 'tab-home';
const BASE_OVERRIDE_VALUE = '__base__';
const LIST_INITIAL_ROWS = 80;
const LIST_ROW_STEP = 60;

const SANITIZE_OPTS = {
  ALLOWED_TAGS: ['p', 'br', 'a', 'strong', 'em', 'u', 'ul', 'ol', 'li', 'blockquote', 'hr', 'img', 'h1', 'h2', 'h3', 'div', 'span'],
  ALLOWED_ATTR: ['href', 'src', 'alt', 'title', 'target', 'rel']
};

const CLUSTER_COLOR_PALETTE = ['#B8952E', '#7B4BA6', '#A8348A', '#2A7B8A', '#4A9B3A', '#9B5A2A', '#C0533F', '#3F6FC0'];
function nextClusterColor(existingCount) {
  return CLUSTER_COLOR_PALETTE[existingCount % CLUSTER_COLOR_PALETTE.length];
}

const dom = {
  body: document.body,
  tabStrip: document.getElementById('tab-strip'),
  bookmarkBar: document.getElementById('bookmark-bar'),
  viewHost: document.getElementById('view-host'),
  tooltip: document.getElementById('floating-tooltip'),
  contextMenu: document.getElementById('bookmark-context-menu'),
  settingsOverlay: document.getElementById('settings-overlay'),
  settingsForm: document.getElementById('settings-form'),
  settingsStatus: document.getElementById('settings-status'),
  clusterEditorOverlay: document.getElementById('cluster-editor-overlay'),
  clusterEditorForm: document.getElementById('cluster-editor-form'),
  clusterEditorTitle: document.getElementById('cluster-editor-title'),
  clusterEditorCopy: document.getElementById('cluster-editor-copy'),
  clusterEditorBookmarkId: document.getElementById('cluster-editor-bookmark-id'),
  clusterEditorName: document.getElementById('cluster-editor-name'),
  clusterEditorColor: document.getElementById('cluster-editor-color'),
  clusterEditorDescription: document.getElementById('cluster-editor-description'),
  thresholdOverlay: document.getElementById('cluster-threshold-overlay'),
  thresholdTitle: document.getElementById('cluster-threshold-title'),
  thresholdResults: document.getElementById('cluster-threshold-results'),
  thresholdSlider: document.getElementById('cluster-threshold-slider'),
  thresholdValue: document.getElementById('cluster-threshold-value'),
  thresholdCount: document.getElementById('cluster-threshold-count'),
  thresholdCreate: document.getElementById('cluster-threshold-create'),
  statusLine: document.getElementById('status-line'),
  statusMsgs: document.getElementById('status-msgs'),
  statusSync: document.getElementById('status-sync'),
  statusHost: document.getElementById('status-host'),
  statusClock: document.getElementById('status-clock'),
  globalStatusLine: document.getElementById('global-status-line'),
  titlebarSession: document.getElementById('titlebar-session')
};

let imapHostLabel = '';

const mailboxShortcuts = [
  { id: 'home', label: 'Home', icon: 'home' },
  { id: 'mailbox:INBOX', label: 'Inbox', icon: 'inbox' },
  { id: 'mailbox:Sent', label: 'Sent', icon: 'send' },
  { id: 'mailbox:Drafts', label: 'Drafts', icon: 'file' },
  { id: 'mailbox:Trash', label: 'Trash', icon: 'trash' }
];

let uniqueId = 0;
let imapEmails = [];
let pcaPoints = getPcaPoints();
let threadHeadersLoaded = false;
let threadCache = null;
let threadIndexByEmailId = {};
let threadSizesByEmailId = {};
let isFetchingFromImap = false;
let isFetchingMore = false;
let imapInboxHasMore = true;
let pendingPromptCluster = null;
let contextBookmarkId = null;
const searchTimers = {};
const composeSelections = {};
localStorage.removeItem('butter-mail-categories');
let storedEmails = getJson(STORAGE_KEY, []);
let storedEmbeddings = getJson(EMBEDDINGS_KEY, {});
let storedPromptClusters = getJson(PROMPT_CLUSTERS_KEY, {});
let storedBookmarkOverrides = getJson(BOOKMARK_OVERRIDES_KEY, {});

(function backfillPromptClusterColors() {
  const slugs = Object.keys(storedPromptClusters);
  let dirty = false;
  slugs.forEach((slug, idx) => {
    if (!storedPromptClusters[slug].color) {
      storedPromptClusters[slug].color = nextClusterColor(idx);
      dirty = true;
    }
  });
  if (dirty) saveJson(PROMPT_CLUSTERS_KEY, storedPromptClusters);
})();
let allEmailsCache = null;
let allEmailsByIdCache = null;
let bookmarkDefinitionsCache = null;
let systemEmailCache = {};
let emailCollectionRevision = 0;
let bookmarkRevision = 0;

const appState = {
  theme: 'dark',
  tabs: [createHomeTab()],
  activeTabId: HOME_TAB_ID,
  status: 'ready.',
  statusTone: 'muted'
};

function getJson(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function saveJson(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

function invalidateEmailCollections() {
  allEmailsCache = null;
  allEmailsByIdCache = null;
  systemEmailCache = {};
  bookmarkDefinitionsCache = null;
  emailCollectionRevision += 1;
}

function invalidateBookmarkDefinitions() {
  bookmarkDefinitionsCache = null;
  bookmarkRevision += 1;
}

function getEmails() { return storedEmails; }
function saveEmails(emails) {
  storedEmails = Array.isArray(emails) ? emails : [];
  saveJson(STORAGE_KEY, storedEmails);
  invalidateEmailCollections();
}
function getEmbeddings() { return storedEmbeddings; }
function saveEmbeddings(embeddings) {
  storedEmbeddings = embeddings || {};
  saveJson(EMBEDDINGS_KEY, storedEmbeddings);
}
function getPcaModel() { return getJson(PCA_KEY, null); }
function savePcaModel(model) { if (model) saveJson(PCA_KEY, model); else localStorage.removeItem(PCA_KEY); }
function getPcaPoints() { return getJson(PCA_POINTS_KEY, {}); }
function savePcaPoints(points) { if (points && Object.keys(points).length) saveJson(PCA_POINTS_KEY, points); else localStorage.removeItem(PCA_POINTS_KEY); }
function getPromptClusters() { return storedPromptClusters; }
function savePromptClusters(clusters) {
  storedPromptClusters = clusters || {};
  saveJson(PROMPT_CLUSTERS_KEY, storedPromptClusters);
  invalidateBookmarkDefinitions();
}
function getBookmarkOverrides() { return storedBookmarkOverrides; }
function saveBookmarkOverrides(overrides) {
  storedBookmarkOverrides = overrides || {};
  saveJson(BOOKMARK_OVERRIDES_KEY, storedBookmarkOverrides);
  invalidateBookmarkDefinitions();
}

function openImapCacheDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IMAP_CACHE_DB_NAME, 1);
    req.onerror = () => reject(req.error);
    req.onsuccess = () => resolve(req.result);
    req.onupgradeneeded = (event) => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains(IMAP_CACHE_STORE)) {
        db.createObjectStore(IMAP_CACHE_STORE, { keyPath: 'accountKey' });
      }
    };
  });
}

async function getCachedEmails(accountKey) {
  if (!accountKey) return [];
  try {
    const db = await openImapCacheDB();
    return await new Promise((resolve, reject) => {
      const req = db.transaction(IMAP_CACHE_STORE, 'readonly').objectStore(IMAP_CACHE_STORE).get(accountKey);
      req.onsuccess = () => resolve(req.result && Array.isArray(req.result.emails) ? req.result.emails : []);
      req.onerror = () => reject(req.error);
    });
  } catch {
    return [];
  }
}

async function setCachedEmails(accountKey, emails) {
  if (!accountKey) return;
  try {
    const db = await openImapCacheDB();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IMAP_CACHE_STORE, 'readwrite');
      tx.objectStore(IMAP_CACHE_STORE).put({ accountKey, emails, lastSynced: Date.now() });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('[butter-mail] cache write failed', err);
  }
}

function createHomeTab() {
  return { id: HOME_TAB_ID, type: 'home', title: 'Home', iconName: 'home', closable: false, query: '', searchResults: null, searchLoading: false };
}

function createClusterTab(options) {
  return {
    id: makeId('tab-cluster'),
    type: 'clusterList',
    sourceType: options.sourceType,
    bookmarkId: options.bookmarkId || '',
    systemId: options.systemId || '',
    title: options.title,
    iconName: options.iconName || 'bookmark',
    closable: true,
    viewMode: 'list',
    query: '',
    searchResults: null,
    searchLoading: false,
    selectedIds: [],
    visibleRows: LIST_INITIAL_ROWS
  };
}

function createEmailTab(emailId) {
  const email = getEmailById(emailId);
  return {
    id: makeId('tab-email'),
    type: 'emailThread',
    emailId,
    title: truncate(email && email.subject ? email.subject : '(no subject)', 34),
    iconName: 'mail-open',
    closable: true,
    expandedMessageIds: { [emailId]: true }
  };
}

function createComposeTab(seed) {
  const tab = {
    id: makeId('tab-compose'),
    type: 'compose',
    title: 'New Email',
    iconName: 'compose',
    closable: true,
    to: seed && seed.to ? seed.to : '',
    subject: seed && seed.subject ? seed.subject : '',
    bodyHtml: seed && seed.bodyHtml ? seed.bodyHtml : '',
    attachments: seed && Array.isArray(seed.attachments) ? seed.attachments : [],
    status: '',
    statusTone: 'muted',
    sending: false,
    replyToMessageId: seed && seed.replyToMessageId ? seed.replyToMessageId : ''
  };
  updateComposeTitle(tab);
  return tab;
}

function makeId(prefix) { uniqueId += 1; return prefix + '-' + Date.now() + '-' + uniqueId; }

function escapeHtml(value) {
  if (value == null) return '';
  const div = document.createElement('div');
  div.textContent = String(value);
  return div.innerHTML;
}

function truncate(value, max) {
  const text = String(value || '');
  return text.length <= max ? text : text.slice(0, Math.max(0, max - 1)).trimEnd() + '…';
}

function formatDate(value) {
  if (!value) return '';
  try {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return date.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: date.getFullYear() !== new Date().getFullYear() ? 'numeric' : undefined });
  } catch {
    return String(value);
  }
}

function formatDateTime(value) {
  if (!value) return '';
  try {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return date.toLocaleString();
  } catch {
    return String(value);
  }
}

function stripHtml(value) {
  return String(value || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function linkify(text) {
  return escapeHtml(text).replace(/(https?:\/\/[^\s<]+)/g, (match) => '<a href="' + escapeHtml(match) + '" target="_blank" rel="noopener">' + escapeHtml(match) + '</a>').replace(/\n/g, '<br>');
}

function renderEmailBody(body, bodyIsHtml) {
  if (!body) return '';
  if (typeof DOMPurify === 'undefined') return escapeHtml(body).replace(/\n/g, '<br>');
  return DOMPurify.sanitize(bodyIsHtml ? body : linkify(body), SANITIZE_OPTS);
}

function sanitizeComposeHtml(html) {
  if (!html) return '';
  return typeof DOMPurify === 'undefined' ? html : DOMPurify.sanitize(html, SANITIZE_OPTS);
}

const ICON_GLYPHS = {
  home: '~',
  inbox: '@',
  send: '>',
  file: '#',
  trash: 'x',
  bookmark: '*',
  sparkles: '+',
  mail: '@',
  'mail-open': '@',
  compose: '!',
  settings: '%',
  plus: '+',
  close: 'x',
  search: '/',
  refresh: '^',
  cluster: '#',
  reply: '<',
  forward: '>',
  chart: '#',
  list: '=',
  attach: '&'
};

function iconMarkup(name, className) {
  const glyph = ICON_GLYPHS[name] || ICON_GLYPHS.mail;
  const classes = className ? escapeHtml(className) : 'glyph';
  return '<span class="' + classes + ' glyph" aria-hidden="true">' + glyph + '</span>';
}

function iconLabelMarkup(iconName, label, options) {
  const opts = options || {};
  const wrapperClass = opts.wrapperClass || 'icon-label';
  const iconClass = opts.iconClass || 'glyph';
  const labelClass = opts.labelClass || '';
  return '<span class="' + escapeHtml(wrapperClass) + '">' +
    iconMarkup(iconName, iconClass) +
    (label ? '<span' + (labelClass ? ' class="' + escapeHtml(labelClass) + '"' : '') + '>' + escapeHtml(label) + '</span>' : '') +
  '</span>';
}

function getAllEmails() {
  if (!allEmailsCache) allEmailsCache = [...imapEmails, ...getEmails()];
  return allEmailsCache;
}

function getEmailById(emailId) {
  if (!allEmailsByIdCache) {
    allEmailsByIdCache = {};
    getAllEmails().forEach((email) => { allEmailsByIdCache[email.id] = email; });
  }
  return allEmailsByIdCache[emailId] || null;
}

function getActiveTab() { return appState.tabs.find((tab) => tab.id === appState.activeTabId) || appState.tabs[0]; }
function updateComposeTitle(tab) { tab.title = truncate(tab.to || 'New Email', 28); }

function setStatus(message, tone) {
  appState.status = message || '';
  appState.statusTone = tone || 'muted';
  const status = document.getElementById('global-status');
  if (status) {
    status.textContent = appState.status;
    status.className = 'status-copy ' + appState.statusTone;
  }
  if (dom.globalStatusLine) {
    dom.globalStatusLine.textContent = appState.status || 'ready.';
    dom.globalStatusLine.className = 'status-seg status-tone-' + (appState.statusTone || 'muted');
  }
}

setInterval(() => updateStatusLine(), 30000);

function pad2(n) { return n < 10 ? '0' + n : String(n); }

function updateStatusLine() {
  if (!dom.statusLine) return;
  const inboxCount = getAllEmails().filter((email) => normalizeMailbox(email) === 'inbox').length;
  if (dom.statusMsgs) dom.statusMsgs.textContent = String(inboxCount) + ' msgs';
  if (dom.statusSync) {
    let sync = 'idle';
    if (isFetchingFromImap) sync = 'sync…';
    else if (isFetchingMore) sync = 'load…';
    dom.statusSync.textContent = sync;
  }
  if (dom.statusHost) dom.statusHost.textContent = 'imap: ' + (imapHostLabel || '—');
  if (dom.statusClock) {
    const now = new Date();
    dom.statusClock.textContent = pad2(now.getHours()) + ':' + pad2(now.getMinutes());
  }
  if (dom.titlebarSession) {
    dom.titlebarSession.textContent = imapHostLabel || 'local';
  }
}

function invalidateThreadCache() {
  threadCache = null;
  threadIndexByEmailId = {};
  threadSizesByEmailId = {};
}

function buildThreadCache() {
  const emails = getAllEmails();
  threadCache = window.ThreadView && typeof window.ThreadView.buildThreads === 'function' ? window.ThreadView.buildThreads(emails) : emails.map((email) => [email]);
  threadIndexByEmailId = {};
  threadSizesByEmailId = {};
  threadCache.forEach((thread) => {
    const size = Array.isArray(thread) ? thread.length : 0;
    (thread || []).forEach((email) => {
      threadIndexByEmailId[email.id] = thread;
      threadSizesByEmailId[email.id] = size;
    });
  });
}

function ensureThreadCacheBuilt() {
  if (!threadCache) buildThreadCache();
}

function sortEmails(emails, searchMode) {
  const list = [...emails];
  if (searchMode) list.sort((left, right) => (left.searchRank || 0) - (right.searchRank || 0));
  else list.sort((left, right) => new Date(right.date || 0) - new Date(left.date || 0));
  return list;
}

function setActiveTab(tabId) { appState.activeTabId = tabId; renderApp(); }

function getTabKey(tab) {
  if (tab.type === 'home') return HOME_TAB_ID;
  if (tab.type === 'clusterList' && tab.sourceType === 'bookmark') return 'bookmark:' + tab.bookmarkId;
  if (tab.type === 'clusterList' && tab.sourceType === 'system') return 'system:' + tab.systemId;
  if (tab.type === 'emailThread') return 'email:' + tab.emailId;
  return '';
}

function addTab(tab, key) {
  const existing = key ? appState.tabs.find((item) => getTabKey(item) === key) : null;
  if (existing) {
    appState.activeTabId = existing.id;
    renderApp();
    return existing;
  }
  appState.tabs.push(tab);
  appState.activeTabId = tab.id;
  renderApp();
  return tab;
}

function closeTab(tabId) {
  const index = appState.tabs.findIndex((tab) => tab.id === tabId);
  if (index === -1 || !appState.tabs[index].closable) return;
  appState.tabs.splice(index, 1);
  if (appState.activeTabId === tabId) {
    const next = appState.tabs[index - 1] || appState.tabs[index] || appState.tabs[0];
    appState.activeTabId = next ? next.id : HOME_TAB_ID;
  }
  renderApp();
}

function openHomeTab() { appState.activeTabId = HOME_TAB_ID; renderApp(); }

function openSystemTab(systemId, label, iconName) {
  return addTab(createClusterTab({ sourceType: 'system', systemId, title: label, iconName: iconName || 'inbox' }), 'system:' + systemId);
}

function openInboxTab() { return openSystemTab('mailbox:INBOX', 'Inbox', 'inbox'); }

function openEmailTab(emailId) {
  const email = getEmailById(emailId);
  if (!email) return;
  addTab(createEmailTab(emailId), 'email:' + emailId);
  ensureEmailBodyLoaded(emailId);
}

function openComposeTab(seed) {
  appState.tabs.push(createComposeTab(seed || {}));
  appState.activeTabId = appState.tabs[appState.tabs.length - 1].id;
  renderApp();
}

function renderApp(options) {
  const next = options || {};
  cleanupOverrides();
  if (next.tabs !== false) renderTabStrip();
  if (next.bookmarks !== false) renderBookmarkBar();
  if (next.view !== false) renderActiveView();
  updateStatusLine();
}

function sortedPromptClusters() {
  return Object.entries(getPromptClusters()).sort((left, right) => {
    const orderA = typeof left[1].order === 'number' ? left[1].order : Number.MAX_SAFE_INTEGER;
    const orderB = typeof right[1].order === 'number' ? right[1].order : Number.MAX_SAFE_INTEGER;
    if (orderA !== orderB) return orderA - orderB;
    const createdA = left[1].createdAt || '';
    const createdB = right[1].createdAt || '';
    if (createdA !== createdB) return createdA.localeCompare(createdB);
    return left[0].localeCompare(right[0]);
  });
}

function getPromptClusterMemberIds(cluster) {
  const ids = new Set(Array.isArray(cluster.emailIds) ? cluster.emailIds : []);
  if (Array.isArray(cluster.scored)) {
    const threshold = cluster.threshold != null ? cluster.threshold : 0.3;
    const overrides = cluster.overrides || {};
    cluster.scored.forEach((item) => {
      const included = overrides[item.id] === true || (overrides[item.id] !== false && item.sim >= threshold);
      if (included) ids.add(item.id);
    });
  }
  return ids;
}

function emailMatchesBookmark(email, bookmarkId, promptMemberIds) {
  const override = getBookmarkOverrides()[email.id];
  if (override) return override === bookmarkId;
  if (bookmarkId.startsWith('prompt:')) {
    const cluster = getPromptClusters()[bookmarkId.slice(7)];
    if (!cluster) return false;
    const ids = promptMemberIds || getPromptClusterMemberIds(cluster);
    return ids.has(email.id);
  }
  return false;
}

function countEmailsForBookmark(bookmarkId, promptMemberIds) {
  return getAllEmails().filter((email) => emailMatchesBookmark(email, bookmarkId, promptMemberIds)).length;
}

function getBookmarkDefinitions() {
  if (bookmarkDefinitionsCache) return bookmarkDefinitionsCache;
  const bookmarks = [];
  sortedPromptClusters().forEach(([slug, cluster]) => {
    const memberIds = getPromptClusterMemberIds(cluster);
    bookmarks.push({
      id: 'prompt:' + slug,
      slug,
      label: cluster.label || slug,
      description: cluster.description || '',
      color: cluster.color || '#7c83ff',
      kind: 'user',
      count: countEmailsForBookmark('prompt:' + slug, memberIds)
    });
  });
  bookmarkDefinitionsCache = bookmarks;
  return bookmarkDefinitionsCache;
}

function getBookmarkById(bookmarkId) {
  return getBookmarkDefinitions().find((bookmark) => bookmark.id === bookmarkId) || null;
}

function openBookmarkTab(bookmarkId) {
  const bookmark = getBookmarkById(bookmarkId);
  if (!bookmark) return;
  addTab(createClusterTab({
    sourceType: 'bookmark',
    bookmarkId,
    title: bookmark.label,
    iconName: 'bookmark'
  }), 'bookmark:' + bookmarkId);
}

function cleanupOverrides() {
  const valid = new Set(getBookmarkDefinitions().map((bookmark) => bookmark.id));
  const overrides = getBookmarkOverrides();
  let dirty = false;
  Object.keys(overrides).forEach((emailId) => {
    if (!valid.has(overrides[emailId])) {
      delete overrides[emailId];
      dirty = true;
    }
  });
  if (dirty) saveBookmarkOverrides(overrides);
}

function getEmailPreview(email) {
  const body = email.bodyIsHtml ? stripHtml(email.body) : String(email.body || '');
  return truncate(body || '(no preview available)', 120);
}

function normalizeMailbox(email) {
  const mailbox = String(email && email.mailbox ? email.mailbox : '').toLowerCase();
  if (email && email.isSent) return 'sent';
  if (mailbox === 'inbox') return 'inbox';
  if (mailbox.includes('sent')) return 'sent';
  if (mailbox.includes('trash') || mailbox.includes('bin') || mailbox.includes('deleted')) return 'trash';
  if (mailbox.includes('draft')) return 'drafts';
  return mailbox;
}

function getEmailsForSystemTab(tab) {
  const key = tab.systemId || 'mailbox:INBOX';
  if (systemEmailCache[key]) return systemEmailCache[key];
  const emails = getAllEmails();
  let filtered = emails.filter((email) => normalizeMailbox(email) === 'inbox');
  if (tab.systemId === 'mailbox:INBOX') filtered = emails.filter((email) => normalizeMailbox(email) === 'inbox');
  else if (tab.systemId === 'mailbox:Sent') filtered = emails.filter((email) => normalizeMailbox(email) === 'sent');
  else if (tab.systemId === 'mailbox:Drafts') filtered = emails.filter((email) => normalizeMailbox(email) === 'drafts');
  else if (tab.systemId === 'mailbox:Trash') filtered = emails.filter((email) => normalizeMailbox(email) === 'trash');
  systemEmailCache[key] = filtered;
  return filtered;
}

function getRenderedEmailCount(tab, total) {
  const limit = tab && tab.type === 'clusterList' ? Math.max(LIST_INITIAL_ROWS, Number(tab.visibleRows) || LIST_INITIAL_ROWS) : total;
  return Math.min(total, limit);
}

function getEmailsForTab(tab) {
  const cache = tab && tab._emailsCache;
  if (cache &&
    cache.emailRevision === emailCollectionRevision &&
    cache.bookmarkRevision === bookmarkRevision &&
    cache.query === String(tab.query || '') &&
    cache.searchResults === tab.searchResults) {
    return cache.value;
  }
  let base = [];
  if (tab.type === 'home') base = getAllEmails();
  else if (tab.type === 'clusterList' && tab.sourceType === 'bookmark') base = getAllEmails().filter((email) => emailMatchesBookmark(email, tab.bookmarkId));
  else if (tab.type === 'clusterList' && tab.sourceType === 'system') base = getEmailsForSystemTab(tab);
  else if (tab.type === 'emailThread') {
    const email = getEmailById(tab.emailId);
    base = email ? [email] : [];
  }
  let value = [];
  if (!tab.query || !tab.query.trim()) value = sortEmails(base, false);
  else if (Array.isArray(tab.searchResults)) {
    const allowed = new Set(base.map((email) => email.id));
    value = sortEmails(tab.searchResults.filter((email) => allowed.has(email.id)), true);
  } else value = sortEmails(base, false);
  if (tab) {
    tab._emailsCache = {
      emailRevision: emailCollectionRevision,
      bookmarkRevision,
      query: String(tab.query || ''),
      searchResults: tab.searchResults,
      value
    };
  }
  return value;
}

function renderTabStrip() {
  dom.tabStrip.innerHTML = appState.tabs.map((tab, index) => {
    const active = tab.id === appState.activeTabId ? ' active' : '';
    const num = String(index + 1);
    return '<div class="browser-tab' + active + '" data-action="activate-tab" data-tab-id="' + escapeHtml(tab.id) + '" role="button" tabindex="0">' +
      '<span class="tab-num">' + num + '</span>' +
      '<span class="tab-glyph">' + (ICON_GLYPHS[tab.iconName] || '@') + '</span>' +
      '<span class="browser-tab-title">' + escapeHtml(tab.title.toLowerCase()) + '</span>' +
      (tab.closable ? '<button type="button" class="tab-close" data-action="close-tab" data-tab-id="' + escapeHtml(tab.id) + '" aria-label="Close tab">x</button>' : '') +
    '</div>';
  }).join('');
}

function renderBookmarkBar() {
  const active = getActiveTab();
  const currentBookmarkId = active && active.type === 'clusterList' && active.sourceType === 'bookmark' ? active.bookmarkId : '';
  dom.bookmarkBar.innerHTML = getBookmarkDefinitions().map((bookmark) => {
    const tooltip = 'Smart cluster' + (bookmark.description ? '\n' + bookmark.description : '');
    const activeClass = bookmark.id === currentBookmarkId ? ' active' : '';
    return '<button type="button" class="bookmark-pill' + activeClass + '" data-action="open-bookmark-tab" data-bookmark-id="' + escapeHtml(bookmark.id) + '" data-tooltip="' + escapeHtml(tooltip) + '">' +
      '<span class="bookmark-swatch" style="background:' + escapeHtml(bookmark.color || '#7c83ff') + '"></span>' +
      '<span>' + escapeHtml(bookmark.label.toLowerCase()) + '</span>' +
      '<span class="bookmark-count" style="color:var(--fg-mute);font-size:11px;">(' + String(bookmark.count) + ')</span>' +
    '</button>';
  }).join('');
}

function renderBookmarkCard(bookmark) {
  const color = bookmark.color || '#7c83ff';
  return '<button type="button" class="bookmark-card" data-action="open-bookmark-tab" data-bookmark-id="' + escapeHtml(bookmark.id) + '" data-tooltip="' + escapeHtml('smart cluster' + (bookmark.description ? '\n' + bookmark.description : '')) + '">' +
    '<div class="bookmark-card-header">' +
      '<h3 class="bookmark-card-name"><span class="bookmark-swatch" style="background:' + escapeHtml(color) + '"></span>' + escapeHtml(bookmark.label.toLowerCase()) + '</h3>' +
      '<span class="bookmark-card-count">' + String(bookmark.count) + '</span>' +
    '</div>' +
    '<p class="bookmark-card-description">' + escapeHtml(bookmark.description || 'smart cluster.') + '</p>' +
  '</button>';
}

function renderSearchPreviewList(results) {
  if (!results.length) return '<div class="state-block">No results yet.</div>';
  return '<div class="search-preview-list">' + results.map((email) => {
    return '<div class="search-preview-row" data-action="open-email-tab" data-email-id="' + escapeHtml(email.id) + '">' +
      '<div class="row-subject">' + escapeHtml(email.subject || '(no subject)') + '</div>' +
      '<div class="row-sender">' + escapeHtml(email.from || email.fromEmail || '') + ' · ' + escapeHtml(formatDate(email.date)) + '</div>' +
      '<div class="row-preview">' + escapeHtml(getEmailPreview(email)) + '</div>' +
    '</div>';
  }).join('') + '</div>';
}

function renderActiveView() {
  const tab = getActiveTab();
  if (!tab) return;
  if (tab.type === 'home') dom.viewHost.innerHTML = renderHomeView(tab);
  else if (tab.type === 'clusterList') dom.viewHost.innerHTML = renderClusterListView(tab);
  else if (tab.type === 'emailThread') dom.viewHost.innerHTML = renderEmailView(tab);
  else if (tab.type === 'compose') dom.viewHost.innerHTML = renderComposeView(tab);
  requestAnimationFrame(() => {
    if (tab.type === 'clusterList' && tab.viewMode === 'graph') renderGraphForTab(tab);
  });
}

function renderHomeView(tab) {
  const bookmarks = getBookmarkDefinitions();
  const searchResults = tab.query && Array.isArray(tab.searchResults) ? tab.searchResults.slice(0, 8) : [];
  const totalEmails = getAllEmails().length;
  const embeddingCount = Object.keys(getEmbeddings() || {}).length;
  const inboxCount = getAllEmails().filter((email) => normalizeMailbox(email) === 'inbox').length;
  const sentCount = getAllEmails().filter((email) => normalizeMailbox(email) === 'sent').length;

  const session = imapHostLabel || 'local';
  const motd =
    '  butter-mail v0.1\n' +
    '  session: ' + session + '\n' +
    '  ' + String(totalEmails) + ' cached · ' + String(inboxCount) + ' inbox · ' + String(sentCount) + ' sent · ' +
    String(bookmarks.length) + ' clusters · ' + String(embeddingCount) + ' embeddings';

  const cmd = (action, attrs, name, desc) =>
    '<button type="button" class="home-cmd" data-action="' + action + '"' + (attrs || '') + '>' +
      '<span class="home-cmd-name">' + escapeHtml(name) + '</span>' +
      '<span class="home-cmd-desc">' + escapeHtml(desc) + '</span>' +
    '</button>';

  const commands =
    cmd('open-compose-tab', '', 'compose', 'open a new compose buffer') +
    cmd('open-mailbox-tab', ' data-system-id="mailbox:INBOX" data-label="Inbox" data-icon="inbox"', 'inbox', 'open inbox (' + String(inboxCount) + ')') +
    cmd('open-mailbox-tab', ' data-system-id="mailbox:Sent" data-label="Sent" data-icon="send"', 'sent', 'view sent mail') +
    cmd('open-mailbox-tab', ' data-system-id="mailbox:Drafts" data-label="Drafts" data-icon="file"', 'drafts', 'view drafts') +
    cmd('open-mailbox-tab', ' data-system-id="mailbox:Trash" data-label="Trash" data-icon="trash"', 'trash', 'view trash') +
    cmd('refresh-imap', '', 'refresh', 'pull latest from imap') +
    cmd('compute-embeddings', '', 'embed', 'compute embeddings for all cached mail') +
    cmd('focus-smart-cluster', '', 'smart', 'create a smart cluster from a prompt') +
    cmd('open-settings', '', 'settings', 'configure imap host / credentials');

  const clusterLine = bookmarks.length
    ? '<div class="home-cluster-line">' +
      bookmarks.slice(0, 24).map((bookmark) =>
        '<button type="button" class="home-cluster-chip" data-action="open-bookmark-tab" data-bookmark-id="' + escapeHtml(bookmark.id) + '">' +
          '<span class="bookmark-swatch" style="background:' + escapeHtml(bookmark.color || '#7c83ff') + '"></span>' +
          escapeHtml(bookmark.label.toLowerCase()) +
          ' <span class="home-cluster-count">(' + String(bookmark.count) + ')</span>' +
        '</button>'
      ).join('') +
      '<button type="button" class="home-cluster-chip home-cluster-add" data-action="open-cluster-editor">+ new</button>' +
    '</div>'
    : '<p class="home-dim">no clusters yet — type a prompt below to create one.</p>';

  const searchBlock = tab.query
    ? '<section class="home-section">' +
        '<div class="home-section-label">search ' + escapeHtml(tab.query) + ' →</div>' +
        (tab.searchLoading
          ? '<p class="home-dim">searching…</p>'
          : (searchResults.length
              ? renderSearchPreviewList(searchResults)
              : '<p class="home-dim">no matches.</p>'))
      + '</section>'
    : '';

  return '<section class="tab-view home-view">' +
    '<div class="home-shell">' +
      '<pre class="home-motd">' + escapeHtml(motd) + '</pre>' +

      '<div class="home-prompt">' +
        '<span class="home-prompt-glyph">/</span>' +
        '<input type="text" class="home-prompt-input" data-role="tab-query" data-tab-id="' + escapeHtml(tab.id) + '" placeholder="search your spread…" value="' + escapeHtml(tab.query || '') + '" autofocus />' +
      '</div>' +

      searchBlock +

      '<section class="home-section">' +
        '<div class="home-section-label">commands</div>' +
        '<div class="home-cmd-grid">' + commands + '</div>' +
      '</section>' +

      '<section class="home-section">' +
        '<div class="home-section-label">clusters</div>' +
        clusterLine +
      '</section>' +

      '<section class="home-section">' +
        '<div class="home-section-label">smart cluster</div>' +
        '<form class="home-prompt-form" id="prompt-cluster-form">' +
          '<span class="home-prompt-glyph">&gt;</span>' +
          '<input type="text" id="prompt-cluster-input" placeholder="e.g. invoices, job hunt, design reviews" />' +
          '<button type="submit" class="home-cmd-go">create</button>' +
        '</form>' +
      '</section>' +

      '<p class="home-status status-copy ' + escapeHtml(appState.statusTone) + '" id="global-status">' + escapeHtml(appState.status || 'ready.') + '</p>' +
    '</div>' +
  '</section>';
}

function renderBookmarkOptions(includeBase) {
  const bookmarks = getBookmarkDefinitions();
  return '<option value="">Move to…</option>' +
    (includeBase ? '<option value="' + BASE_OVERRIDE_VALUE + '">Base classification</option>' : '') +
    bookmarks.map((bookmark) => '<option value="' + escapeHtml(bookmark.id) + '">' + escapeHtml(bookmark.label) + '</option>').join('');
}

function renderBatchToolbar(tab) {
  return '<div class="cluster-batch-toolbar"><span>› ' + String(tab.selectedIds.length) + ' selected</span><select data-role="batch-move" data-tab-id="' + escapeHtml(tab.id) + '">' + renderBookmarkOptions(true) + '</select><button type="button" class="secondary-btn" data-action="clear-selection" data-tab-id="' + escapeHtml(tab.id) + '">clear</button></div>';
}

function renderGraphShell(emails) {
  if (!Object.keys(getEmbeddings()).length || !Object.keys(pcaPoints || {}).length) return '<div class="state-block">compute embeddings first to use graph mode.</div>';
  if (!emails.length) return '<div class="state-block">nothing to graph in this tab.</div>';
  return '<div class="graph-shell"><div class="graph-container" id="graph-container"><canvas id="graph-canvas"></canvas><div class="graph-axis-legend"><span>x</span><span>y</span><span>z</span></div><div class="graph-coords-legend" id="graph-coords">x: — y: — z: —</div><div class="graph-tooltip hidden" id="graph-tooltip"></div></div></div>';
}

function renderClusterRows(tab, emails, options) {
  if (!emails.length) return '<div class="state-block">no emails here yet.</div>';
  ensureThreadCacheBuilt();
  const selectable = !options || options.selectable !== false;
  const renderedCount = getRenderedEmailCount(tab, emails.length);
  const visibleEmails = emails.slice(0, renderedCount);
  const remainingCount = Math.max(0, emails.length - renderedCount);
  return '<div class="cluster-list-body"' + (tab.type === 'clusterList' ? ' data-tab-id="' + escapeHtml(tab.id) + '"' : '') + '>' + visibleEmails.map((email) => {
    const selected = tab.selectedIds.includes(email.id) ? ' checked' : '';
    const threadSize = threadSizesByEmailId[email.id] || 1;
    const sender = email.from || email.fromEmail || '(unknown)';
    const mailbox = normalizeMailbox(email);
    return '<div class="cluster-row" data-email-id="' + escapeHtml(email.id) + '">' +
      (selectable ? '<input class="cluster-row-checkbox" type="checkbox" data-role="row-select" data-tab-id="' + escapeHtml(tab.id) + '" data-email-id="' + escapeHtml(email.id) + '"' + selected + ' />' : '<span></span>') +
      '<div class="row-main" data-action="open-email-tab" data-email-id="' + escapeHtml(email.id) + '">' +
        '<div class="row-meta-line">' +
          '<span class="row-identity"><span class="row-sender">' + escapeHtml(sender) + '</span></span>' +
          '<span class="row-separator">·</span>' +
          '<span class="row-mailbox">' + escapeHtml(mailbox || 'mail') + '</span>' +
          (threadSize > 1 ? '<span class="thread-indicator">' + String(threadSize) + '</span>' : '<span></span>') +
          '<span class="row-date">' + escapeHtml(formatDate(email.date)) + '</span>' +
        '</div>' +
        '<div class="row-copy">' +
          '<div class="row-subject">' + escapeHtml(email.subject || '(no subject)') + '</div>' +
          '<div class="row-preview">' + escapeHtml(getEmailPreview(email)) + '</div>' +
        '</div>' +
      '</div>' +
      (selectable
        ? '<div class="cluster-row-actions"><select class="cluster-row-move" data-role="row-move" data-email-id="' + escapeHtml(email.id) + '">' + renderBookmarkOptions(true) + '</select><button type="button" class="cluster-row-open" data-action="open-email-tab" data-email-id="' + escapeHtml(email.id) + '">open</button></div>'
        : '<div class="cluster-row-actions cluster-row-actions-static"><button type="button" class="cluster-row-open" data-action="open-email-tab" data-email-id="' + escapeHtml(email.id) + '">open</button></div>') +
    '</div>';
  }).join('') + (remainingCount ? '<button type="button" class="cluster-list-more" data-action="expand-cluster-list" data-tab-id="' + escapeHtml(tab.id) + '">show ' + String(Math.min(LIST_ROW_STEP, remainingCount)) + ' more (' + String(remainingCount) + ' remaining)</button>' : '') + '</div>';
}

function renderClusterListView(tab) {
  const emails = getEmailsForTab(tab);
  const bookmark = tab.sourceType === 'bookmark' ? getBookmarkById(tab.bookmarkId) : null;
  const title = bookmark ? bookmark.label : tab.title;
  const subtitle = bookmark ? ((bookmark.description || 'smart cluster.') + ' ' + String(emails.length) + ' msg(s).') : String(emails.length) + ' msg(s).';
  const canLoadMore = window.electronAPI && tab.systemId === 'mailbox:INBOX';
  const renderedCount = tab.viewMode === 'graph' ? emails.length : getRenderedEmailCount(tab, emails.length);
  return '<section class="tab-view cluster-view">' +
    '<div class="view-header"><div><h1 class="view-title">' + escapeHtml(title.toLowerCase()) + '</h1><p class="view-subtitle">' + escapeHtml(subtitle) + '</p></div><div class="segmented"><button type="button" class="segmented-btn' + (tab.viewMode === 'list' ? ' active' : '') + '" data-action="set-cluster-view" data-tab-id="' + escapeHtml(tab.id) + '" data-view-mode="list">list</button><button type="button" class="segmented-btn' + (tab.viewMode === 'graph' ? ' active' : '') + '" data-action="set-cluster-view" data-tab-id="' + escapeHtml(tab.id) + '" data-view-mode="graph">graph</button></div></div>' +
    '<div class="cluster-list-panel"><div class="cluster-toolbar"><div class="cluster-search-shell"><input type="text" class="search-field" data-role="tab-query" data-tab-id="' + escapeHtml(tab.id) + '" placeholder="filter this tab…" value="' + escapeHtml(tab.query || '') + '" /></div><div class="toolbar-row">' + (canLoadMore ? '<button type="button" class="toolbar-pill" data-action="load-more">load more</button>' : '') + '<button type="button" class="toolbar-pill" data-action="refresh-imap">refresh</button></div></div>' + (tab.selectedIds.length ? renderBatchToolbar(tab) : '') + (tab.searchLoading ? '<p class="toolbar-note">searching…</p>' : '') + (tab.viewMode === 'graph' ? renderGraphShell(emails) : '<p class="toolbar-note">showing ' + String(renderedCount) + ' of ' + String(emails.length) + '</p>' + renderClusterRows(tab, emails)) + '</div>' +
  '</section>';
}

function renderThreadMessageCard(tab, message, active) {
  const expanded = tab.expandedMessageIds[message.id] || active;
  const cls = 'message-card' + (expanded ? ' expanded' : '');
  return '<article class="' + cls + '">' +
    '<div class="message-card-head" data-action="toggle-thread-message" data-tab-id="' + escapeHtml(tab.id) + '" data-email-id="' + escapeHtml(message.id) + '">' +
      '<div>' +
        '<div class="row-subject">' + escapeHtml(message.subject || '(no subject)') + '</div>' +
        '<div class="row-sender" style="color:var(--fg-mute);font-size:12px;">from: ' + escapeHtml(message.from || message.fromEmail || '') + '</div>' +
      '</div>' +
      '<div class="row-date">' + escapeHtml(formatDateTime(message.date)) + '</div>' +
    '</div>' +
    (expanded ? '<div class="message-card-body"><div class="message-body">' + renderEmailBody(message.body || '', !!message.bodyIsHtml) + '</div></div>' : '') +
  '</article>';
}

function renderEmailView(tab) {
  const email = getEmailById(tab.emailId);
  if (!email) return '<section class="tab-view"><div class="state-block">this email is no longer available.</div></section>';
  ensureThreadCacheBuilt();
  const thread = threadIndexByEmailId[email.id] || [email];
  return '<section class="tab-view message-view"><div class="message-shell">' +
    '<div class="message-reader">' +
      '<h1 class="message-subject">' + escapeHtml(email.subject || '(no subject)') + '</h1>' +
      '<div class="message-meta-grid">' +
        '<div class="message-meta-card"><span class="meta-label">from</span>' + escapeHtml(email.from || email.fromEmail || '') + '</div>' +
        '<div class="message-meta-card"><span class="meta-label">to</span>' + escapeHtml(email.toDisplay || email.to || '') + '</div>' +
        '<div class="message-meta-card"><span class="meta-label">date</span>' + escapeHtml(formatDateTime(email.date)) + '</div>' +
      '</div>' +
      '<div class="message-actions">' +
        '<button type="button" class="message-action-btn" data-action="reply-email" data-email-id="' + escapeHtml(email.id) + '">reply</button>' +
        '<button type="button" class="message-action-btn" data-action="forward-email" data-email-id="' + escapeHtml(email.id) + '">forward</button>' +
        '<select class="cluster-row-move" data-role="row-move" data-email-id="' + escapeHtml(email.id) + '">' + renderBookmarkOptions(true) + '</select>' +
        '<button type="button" class="message-action-btn" data-action="delete-email" data-email-id="' + escapeHtml(email.id) + '">delete</button>' +
      '</div>' +
      '<div class="message-thread">' + thread.map((message) => renderThreadMessageCard(tab, message, message.id === email.id)).join('') + '</div>' +
    '</div>' +
    '<aside class="message-sidebar">' +
      '<h2 class="surface-title">thread</h2>' +
      '<p class="surface-copy">' + String(thread.length) + ' message(s).</p>' +
      '<div class="search-preview-list">' + thread.map((message) => '<button type="button" class="search-preview-row" data-action="switch-email-tab-message" data-tab-id="' + escapeHtml(tab.id) + '" data-email-id="' + escapeHtml(message.id) + '"><div class="row-subject">' + escapeHtml(message.subject || '(no subject)') + '</div><div class="row-sender" style="color:var(--fg-mute);font-size:12px;">' + escapeHtml(message.from || message.fromEmail || '') + '</div><div class="row-date">' + escapeHtml(formatDate(message.date)) + '</div></button>').join('') + '</div>' +
    '</aside>' +
  '</div></section>';
}

function renderAttachmentList(tab) {
  if (!Array.isArray(tab.attachments) || !tab.attachments.length) return '<div class="state-block">no attachments.</div>';
  return tab.attachments.map((attachment) => {
    const parts = String(attachment.path || '').split(/[/\\]/);
    const label = attachment.filename || parts[parts.length - 1] || attachment.path || 'attachment';
    return '<div class="compose-attachment-item"><span>' + escapeHtml(label) + '</span><button type="button" class="secondary-btn" data-action="remove-attachment" data-tab-id="' + escapeHtml(tab.id) + '" data-path="' + escapeHtml(attachment.path) + '">remove</button></div>';
  }).join('');
}

function renderComposeView(tab) {
  return '<section class="tab-view compose-view">' +
    '<div class="view-header"><div><h1 class="view-title">' + escapeHtml(tab.title.toLowerCase()) + '</h1><p class="view-subtitle">writing buffer · esc closes overlays</p></div></div>' +
    '<div class="compose-shell">' +
      '<div class="compose-panel">' +
        '<div class="compose-fields">' +
          '<div class="compose-row"><span class="compose-row-label">to</span><input type="text" data-role="compose-to" data-tab-id="' + escapeHtml(tab.id) + '" placeholder="someone@example.com" value="' + escapeHtml(tab.to) + '" /></div>' +
          '<div class="compose-row"><span class="compose-row-label">subject</span><input type="text" data-role="compose-subject" data-tab-id="' + escapeHtml(tab.id) + '" placeholder="(no subject)" value="' + escapeHtml(tab.subject) + '" /></div>' +
        '</div>' +
        '<div class="compose-toolbar">' +
          '<button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="bold">b</button>' +
          '<button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="italic">i</button>' +
          '<button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="underline">u</button>' +
          '<button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="insertUnorderedList">ul</button>' +
          '<button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="insertOrderedList">ol</button>' +
          '<button type="button" class="compose-tool-btn" data-action="compose-link" data-tab-id="' + escapeHtml(tab.id) + '">link</button>' +
          '<button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="unlink">unlink</button>' +
        '</div>' +
        '<div class="compose-editor" contenteditable="true" spellcheck="true" data-role="compose-editor" data-tab-id="' + escapeHtml(tab.id) + '" data-placeholder="write your message…">' + sanitizeComposeHtml(tab.bodyHtml) + '</div>' +
        '<div class="compose-actions">' +
          '<button type="button" class="secondary-btn" data-action="compose-attach" data-tab-id="' + escapeHtml(tab.id) + '">attach</button>' +
          '<button type="button" class="primary-btn" data-action="send-compose" data-tab-id="' + escapeHtml(tab.id) + '">' + (tab.sending ? 'sending…' : 'send') + '</button>' +
        '</div>' +
        '<p class="compose-status ' + escapeHtml(tab.statusTone || 'muted') + '">' + escapeHtml(tab.status || '') + '</p>' +
      '</div>' +
      '<aside class="compose-side">' +
        '<h2 class="surface-title">attachments</h2>' +
        '<p class="surface-copy">minimal formatting. dedicated compose buffer.</p>' +
        '<div class="compose-attachment-list">' + renderAttachmentList(tab) + '</div>' +
      '</aside>' +
    '</div>' +
  '</section>';
}

function getClusterColorsForEmail(emailId) {
  const overrides = getBookmarkOverrides();
  const override = overrides[emailId];
  if (override) {
    if (override.startsWith('prompt:')) {
      const cluster = getPromptClusters()[override.slice(7)];
      return cluster && cluster.color ? [cluster.color] : [];
    }
    return [];
  }
  const colors = [];
  sortedPromptClusters().forEach(([, cluster]) => {
    const ids = getPromptClusterMemberIds(cluster);
    if (ids.has(emailId) && cluster.color) colors.push(cluster.color);
  });
  return colors;
}

function renderGraphForTab(tab) {
  const container = document.getElementById('graph-container');
  if (!container || !window.GraphView) return;
  const emails = getEmailsForTab(tab);
  const embeddings = getEmbeddings();
  const points = {};
  const emailsById = {};
  emails.forEach((email) => {
    if (embeddings[email.id] && pcaPoints[email.id]) {
      points[email.id] = pcaPoints[email.id];
      emailsById[email.id] = { ...email, clusterColors: getClusterColorsForEmail(email.id) };
    }
  });
  if (!Object.keys(points).length) return;
  if (!window._graphInited && typeof window.GraphView.init === 'function') {
    window.GraphView.init('graph-container');
    window.GraphView.animate();
    window._graphInited = true;
  }
  window.GraphView.render(points, emailsById);
}

window.onGraphPointClick = function (emailId) {
  openEmailTab(emailId);
};

async function ensureThreadHeaders() {
  if (typeof window.electronAPI === 'undefined' || threadHeadersLoaded) return;
  const needsHeaders = imapEmails.filter((email) => email.uid && !email.messageId && (!email.mailbox || email.mailbox === 'INBOX'));
  if (!needsHeaders.length) {
    threadHeadersLoaded = true;
    return;
  }
  try {
    const result = await window.electronAPI.imap.fetchThreadHeaders(needsHeaders.map((email) => email.uid));
    if (!result.ok || !result.headers) return;
    Object.keys(result.headers).forEach((uid) => {
      const email = imapEmails.find((item) => String(item.uid) === String(uid));
      const header = result.headers[uid];
      if (email && header) {
        email.messageId = header.messageId || '';
        email.inReplyTo = header.inReplyTo || '';
        email.references = header.references || '';
      }
    });
    threadHeadersLoaded = true;
    invalidateThreadCache();
  } catch (err) {
    console.warn('[butter-mail] thread headers failed', err);
  }
}

async function ensureEmailBodyLoaded(emailId) {
  const email = getEmailById(emailId);
  if (!email || email.body || !email.uid || typeof window.electronAPI === 'undefined') return;
  const result = await window.electronAPI.imap.fetchOne(email.uid);
  if (result.ok && result.email) {
    email.body = result.email.body || '';
    email.bodyIsHtml = !!result.email.bodyIsHtml;
    email.toDisplay = result.email.toDisplay || email.toDisplay || '';
    email.messageId = result.email.messageId || email.messageId || '';
    email.inReplyTo = result.email.inReplyTo || email.inReplyTo || '';
    email.references = result.email.references || email.references || '';
    invalidateThreadCache();
    renderActiveView();
  }
}

async function refreshFromImap() {
  if (typeof window.electronAPI === 'undefined' || isFetchingFromImap) return;
  isFetchingFromImap = true;
  setStatus('Refreshing from IMAP…', 'muted');
  try {
    const config = await window.electronAPI.imap.getConfig();
    const accountKey = config && config.host && config.user ? config.host + '::' + config.user : '';
    if (config && config.host) imapHostLabel = config.host + (config.user ? ' / ' + config.user : '');
    const result = await window.electronAPI.imap.fetch(150);
    if (!result.ok) {
      if (!String(result.error || '').includes('not configured')) alert('IMAP refresh failed: ' + (result.error || 'Unknown error'));
      return;
    }
    imapEmails = Array.isArray(result.emails) ? result.emails : [];
    invalidateEmailCollections();
    threadHeadersLoaded = false;
    invalidateThreadCache();
    imapInboxHasMore = imapEmails.filter((email) => normalizeMailbox(email) === 'inbox').length >= 150;
    if (accountKey) await setCachedEmails(accountKey, imapEmails);
    rerunTabSearches();
    setStatus('Refreshed.', 'success');
    renderApp();
    ensureThreadHeaders().then(() => renderApp({ tabs: false, bookmarks: false }));
  } finally {
    isFetchingFromImap = false;
  }
}

async function loadMoreImapEmails() {
  if (typeof window.electronAPI === 'undefined' || isFetchingMore || !imapInboxHasMore) return;
  const inbox = imapEmails.filter((email) => normalizeMailbox(email) === 'inbox');
  const beforeUid = inbox.reduce((min, email) => (min == null || (email.uid && email.uid < min) ? email.uid : min), null);
  if (beforeUid == null) return;
  isFetchingMore = true;
  try {
    const result = await window.electronAPI.imap.fetchMore(75, beforeUid);
    if (!result.ok) throw new Error(result.error || 'Load more failed');
    imapInboxHasMore = !!result.hasMore;
    const merged = {};
    [...imapEmails, ...(result.emails || [])].forEach((email) => { merged[email.id] = email; });
    imapEmails = Object.values(merged);
    invalidateEmailCollections();
    invalidateThreadCache();
    rerunTabSearches();
    renderApp();
    ensureThreadHeaders().then(() => renderApp({ tabs: false, bookmarks: false }));
  } catch (err) {
    alert('Could not load more: ' + (err.message || String(err)));
  } finally {
    isFetchingMore = false;
  }
}

async function computeEmbeddings() {
  if (typeof window.electronAPI === 'undefined') {
    alert('Compute embeddings requires the Electron app.');
    return;
  }
  let emails = getAllEmails();
  if (!emails.length) {
    alert('No emails available yet.');
    return;
  }
  const needsBodies = emails.filter((email) => email.id.startsWith('imap-') && !email.body);
  if (needsBodies.length) {
    const bodyResult = await window.electronAPI.imap.fetchBodies(needsBodies.map((email) => email.uid));
    if (bodyResult.ok && bodyResult.bodies) {
      imapEmails = imapEmails.map((email) => bodyResult.bodies[email.uid] ? { ...email, body: bodyResult.bodies[email.uid].body, bodyIsHtml: bodyResult.bodies[email.uid].bodyIsHtml } : email);
      invalidateEmailCollections();
      emails = getAllEmails();
    }
  }
  setStatus('Computing embeddings…', 'muted');
  const embeddingResult = await window.electronAPI.embeddings.compute(emails);
  if (!embeddingResult.ok) {
    alert('Embedding computation failed: ' + (embeddingResult.error || 'Unknown error'));
    return;
  }
  saveEmbeddings(embeddingResult.embeddings || {});
  const emailIds = Object.keys(embeddingResult.embeddings || {});
  const pcaResult = await window.electronAPI.embeddings.pca(embeddingResult.embeddings || {}, emailIds);
  if (pcaResult.ok) {
    pcaPoints = pcaResult.points || {};
    savePcaPoints(pcaPoints);
    savePcaModel(pcaResult.model || null);
  }
  setStatus('Embeddings updated.', 'success');
  renderApp();
}

function createClusterSlug(label, clusters) {
  const base = label.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'cluster';
  let slug = base;
  let index = 1;
  while (clusters[slug]) {
    index += 1;
    slug = base + '-' + index;
  }
  return slug;
}

async function createPromptClusterFromPrompt(promptValue) {
  const prompt = String(promptValue || '').trim();
  if (!prompt) return;
  if (typeof window.electronAPI === 'undefined') {
    alert('Prompt clusters require the Electron app.');
    return;
  }
  const embeddings = getEmbeddings();
  const emailIds = Object.keys(embeddings);
  if (!emailIds.length) {
    alert('Compute embeddings first.');
    return;
  }
  setStatus('Finding related emails…', 'muted');
  const result = await window.electronAPI.embeddings.promptClusterScored(prompt, embeddings, emailIds);
  if (!result.ok) {
    alert('Cluster creation failed: ' + (result.error || 'Unknown error'));
    return;
  }
  pendingPromptCluster = { prompt, scored: result.scored || [], overrides: {} };
  dom.thresholdTitle.textContent = 'Create cluster: "' + prompt + '"';
  dom.thresholdSlider.value = '0.30';
  renderThresholdResults();
  dom.thresholdOverlay.classList.remove('hidden');
}

function renderThresholdResults() {
  if (!pendingPromptCluster) return;
  const threshold = parseFloat(dom.thresholdSlider.value || '0.3');
  dom.thresholdValue.textContent = threshold.toFixed(2);
  const emailsById = {};
  getAllEmails().forEach((email) => { emailsById[email.id] = email; });
  let count = 0;
  dom.thresholdResults.innerHTML = (pendingPromptCluster.scored || []).map((item) => {
    const included = pendingPromptCluster.overrides[item.id] === true || (pendingPromptCluster.overrides[item.id] !== false && item.sim >= threshold);
    if (included) count += 1;
    const email = emailsById[item.id];
    return '<div class="cluster-result-row ' + (included ? 'in-cluster' : 'not-in-cluster') + '" data-role="threshold-row" data-email-id="' + escapeHtml(item.id) + '"><span>' + escapeHtml(email && email.subject ? email.subject : '(no subject)') + '</span><span>' + item.sim.toFixed(2) + '</span></div>';
  }).join('') || '<div class="state-block">No matches.</div>';
  dom.thresholdCount.textContent = count + ' email(s) in cluster';
}

function closeThresholdOverlay() {
  dom.thresholdOverlay.classList.add('hidden');
  pendingPromptCluster = null;
}

function commitThresholdCluster() {
  if (!pendingPromptCluster) return;
  const clusters = getPromptClusters();
  const slug = createClusterSlug(pendingPromptCluster.prompt, clusters);
  clusters[slug] = {
    label: pendingPromptCluster.prompt,
    description: '',
    color: nextClusterColor(Object.keys(clusters).length),
    threshold: parseFloat(dom.thresholdSlider.value || '0.3'),
    scored: pendingPromptCluster.scored || [],
    overrides: pendingPromptCluster.overrides || {},
    emailIds: [],
    createdAt: new Date().toISOString(),
    order: Object.keys(clusters).length
  };
  savePromptClusters(clusters);
  closeThresholdOverlay();
  renderApp();
  openBookmarkTab('prompt:' + slug);
}

function scheduleSearch(tab) {
  clearTimeout(searchTimers[tab.id]);
  if (!tab.query || !tab.query.trim()) {
    tab.searchLoading = false;
    tab.searchResults = null;
    renderActiveView();
    return;
  }
  tab.searchLoading = true;
  searchTimers[tab.id] = setTimeout(async () => {
    const query = tab.query.trim();
    let results = [];
    if (typeof window.electronAPI !== 'undefined' && window.HybridSearch) {
      results = await window.HybridSearch.search(query, getAllEmails(), getEmbeddings(), window.electronAPI);
    } else {
      const lower = query.toLowerCase();
      results = getAllEmails().filter((email) => String(email.subject || '').toLowerCase().includes(lower) || String(email.body || '').toLowerCase().includes(lower) || String(email.from || '').toLowerCase().includes(lower) || String(email.fromEmail || '').toLowerCase().includes(lower));
    }
    tab.searchResults = results;
    tab.searchLoading = false;
    if (getActiveTab().id === tab.id) renderActiveView();
  }, 260);
}

function rerunTabSearches() {
  appState.tabs.forEach((tab) => {
    if (tab.query && tab.query.trim()) scheduleSearch(tab);
  });
}

function updateClusterSelection(tabId, emailId, checked) {
  const tab = appState.tabs.find((item) => item.id === tabId && item.type === 'clusterList');
  if (!tab) return;
  const selected = new Set(tab.selectedIds || []);
  if (checked) selected.add(emailId);
  else selected.delete(emailId);
  tab.selectedIds = Array.from(selected);
  renderActiveView();
}

function clearClusterSelection(tabId) {
  const tab = appState.tabs.find((item) => item.id === tabId && item.type === 'clusterList');
  if (!tab) return;
  tab.selectedIds = [];
  renderActiveView();
}

function expandClusterList(tabId) {
  const tab = appState.tabs.find((item) => item.id === tabId && item.type === 'clusterList');
  if (!tab) return;
  tab.visibleRows = Math.max(LIST_INITIAL_ROWS, Number(tab.visibleRows) || LIST_INITIAL_ROWS) + LIST_ROW_STEP;
  if (getActiveTab().id === tab.id) renderActiveView();
}

function moveEmailsToBookmark(emailIds, targetBookmarkId) {
  const ids = Array.isArray(emailIds) ? emailIds.filter(Boolean) : [];
  if (!ids.length) return;
  const overrides = getBookmarkOverrides();
  ids.forEach((emailId) => {
    if (!targetBookmarkId || targetBookmarkId === BASE_OVERRIDE_VALUE) delete overrides[emailId];
    else overrides[emailId] = targetBookmarkId;
  });
  saveBookmarkOverrides(overrides);
  appState.tabs.forEach((tab) => { if (tab.type === 'clusterList') tab.selectedIds = []; });
  renderApp();
}

function saveComposeSelection(tabId) {
  const editor = document.querySelector('[data-role="compose-editor"][data-tab-id="' + CSS.escape(tabId) + '"]');
  if (!editor) return;
  const selection = window.getSelection ? window.getSelection() : null;
  if (!selection || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  const common = range.commonAncestorContainer;
  const parent = common && (common.nodeType === 1 ? common : common.parentElement);
  if (parent && (parent === editor || editor.contains(parent))) composeSelections[tabId] = range.cloneRange();
}

function restoreComposeSelection(tabId) {
  const selection = window.getSelection ? window.getSelection() : null;
  const range = composeSelections[tabId];
  if (!selection || !range) return;
  try {
    selection.removeAllRanges();
    selection.addRange(range);
  } catch (_) {}
}

function updateComposeStateFromEditor(tabId, editor) {
  const tab = appState.tabs.find((item) => item.id === tabId && item.type === 'compose');
  if (!tab || !editor) return;
  tab.bodyHtml = editor.innerHTML || '';
}

function execComposeCommand(tabId, command) {
  const editor = document.querySelector('[data-role="compose-editor"][data-tab-id="' + CSS.escape(tabId) + '"]');
  if (!editor) return;
  editor.focus();
  restoreComposeSelection(tabId);
  try { document.execCommand(command, false, null); } catch (_) {}
  updateComposeStateFromEditor(tabId, editor);
  saveComposeSelection(tabId);
}

function normalizeLinkUrl(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (/^https?:\/\//i.test(text) || /^mailto:/i.test(text)) return text;
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text)) return 'mailto:' + text;
  return 'https://' + text.replace(/^\/*/, '');
}

function insertComposeLink(tabId) {
  const editor = document.querySelector('[data-role="compose-editor"][data-tab-id="' + CSS.escape(tabId) + '"]');
  if (!editor) return;
  const url = normalizeLinkUrl(prompt('Enter a URL or email address'));
  if (!url) return;
  editor.focus();
  restoreComposeSelection(tabId);
  const selection = window.getSelection ? window.getSelection() : null;
  if (selection && selection.rangeCount > 0 && !selection.getRangeAt(0).collapsed) {
    try { document.execCommand('createLink', false, url); } catch (_) {}
  } else {
    try { document.execCommand('insertHTML', false, '<a href="' + escapeHtml(url) + '" target="_blank" rel="noopener">' + escapeHtml(url) + '</a>'); } catch (_) {}
  }
  updateComposeStateFromEditor(tabId, editor);
  saveComposeSelection(tabId);
}

function updateComposeTitleFromField(tabId) {
  const tab = appState.tabs.find((item) => item.id === tabId && item.type === 'compose');
  if (!tab) return;
  updateComposeTitle(tab);
  renderTabStrip();
}

function removeEmailFromState(emailId) {
  imapEmails = imapEmails.filter((email) => email.id !== emailId);
  saveEmails(getEmails().filter((email) => email.id !== emailId));
  const overrides = getBookmarkOverrides();
  delete overrides[emailId];
  saveBookmarkOverrides(overrides);
  appState.tabs = appState.tabs.filter((tab) => !(tab.type === 'emailThread' && tab.emailId === emailId));
  appState.tabs.forEach((tab) => { if (tab.type === 'clusterList') tab.selectedIds = (tab.selectedIds || []).filter((id) => id !== emailId); });
  if (!appState.tabs.some((tab) => tab.id === appState.activeTabId)) appState.activeTabId = HOME_TAB_ID;
  invalidateEmailCollections();
  invalidateThreadCache();
}

async function sendCompose(tabId) {
  const tab = appState.tabs.find((item) => item.id === tabId && item.type === 'compose');
  if (!tab || tab.sending) return;
  if (typeof window.electronAPI === 'undefined') {
    alert('Sending is available in the Electron app.');
    return;
  }
  const editor = document.querySelector('[data-role="compose-editor"][data-tab-id="' + CSS.escape(tabId) + '"]');
  if (editor) updateComposeStateFromEditor(tabId, editor);
  tab.sending = true;
  tab.status = 'Sending…';
  tab.statusTone = 'muted';
  renderActiveView();
  const bodyText = stripHtml(tab.bodyHtml);
  const butterMailUrl = 'https://github.com/mehek-niwas/butter-mail';
  const result = await window.electronAPI.smtp.send({
    to: tab.to,
    subject: tab.subject,
    text: (bodyText ? bodyText + '\n\n' : '') + '- sent with butter mail ' + butterMailUrl,
    html: (sanitizeComposeHtml(tab.bodyHtml) || '') + '<br><br>- sent with <a href="' + butterMailUrl + '" target="_blank" rel="noopener">butter mail</a>',
    replyToMessageId: tab.replyToMessageId,
    attachments: tab.attachments
  });
  if (!result.ok) {
    tab.sending = false;
    tab.status = result.error || 'Send failed.';
    tab.statusTone = 'error';
    renderActiveView();
    return;
  }
  tab.status = 'Sent.';
  tab.statusTone = 'success';
  renderActiveView();
  await refreshFromImap();
  setTimeout(() => closeTab(tabId), 700);
}

async function attachFilesToCompose(tabId) {
  const tab = appState.tabs.find((item) => item.id === tabId && item.type === 'compose');
  if (!tab || typeof window.electronAPI === 'undefined' || !window.electronAPI.dialog) return;
  const result = await window.electronAPI.dialog.pickAttachments();
  if (!result.ok || !Array.isArray(result.filePaths)) return;
  const seen = new Set(tab.attachments.map((attachment) => attachment.path));
  result.filePaths.forEach((path) => {
    const next = String(path || '').trim();
    if (next && !seen.has(next)) tab.attachments.push({ path: next });
  });
  renderActiveView();
}

function removeComposeAttachment(tabId, path) {
  const tab = appState.tabs.find((item) => item.id === tabId && item.type === 'compose');
  if (!tab) return;
  tab.attachments = tab.attachments.filter((attachment) => attachment.path !== path);
  renderActiveView();
}

function replyEmail(emailId) {
  const email = getEmailById(emailId);
  if (!email) return;
  const subject = /^re:/i.test(email.subject || '') ? (email.subject || '(no subject)') : ('Re: ' + (email.subject || '(no subject)'));
  openComposeTab({ to: email.fromEmail || email.from || '', subject, replyToMessageId: email.messageId || '' });
}

function forwardEmail(emailId) {
  const email = getEmailById(emailId);
  if (!email) return;
  const forwarded = '\n\n---------- Forwarded message ----------\nFrom: ' + (email.from || email.fromEmail || '') + '\nDate: ' + formatDateTime(email.date) + '\nSubject: ' + (email.subject || '(no subject)') + '\n\n' + stripHtml(email.body || '');
  openComposeTab({ subject: 'Fwd: ' + (email.subject || '(no subject)'), bodyHtml: escapeHtml(forwarded).replace(/\n/g, '<br>') });
}

async function handleDeleteEmail(emailId) {
  const email = getEmailById(emailId);
  if (!email || !confirm('Delete this email?')) return;
  if (email.id.startsWith('imap-') && typeof window.electronAPI !== 'undefined' && window.electronAPI.imap && window.electronAPI.imap.delete) {
    const result = await window.electronAPI.imap.delete({ uid: email.uid, mailbox: email.mailbox || 'INBOX' });
    if (!result.ok) {
      alert('Delete failed: ' + (result.error || 'Unknown error'));
      return;
    }
  }
  removeEmailFromState(emailId);
  renderApp();
}

function openClusterEditor(bookmarkId, focusMode) {
  contextBookmarkId = bookmarkId || '';
  if (!bookmarkId) {
    dom.clusterEditorTitle.textContent = 'Create cluster';
    dom.clusterEditorCopy.textContent = 'Create a smart cluster for email triage.';
    dom.clusterEditorBookmarkId.value = '';
    dom.clusterEditorName.value = '';
    dom.clusterEditorDescription.value = '';
    if (dom.clusterEditorColor) dom.clusterEditorColor.value = nextClusterColor(Object.keys(getPromptClusters()).length);
  } else if (bookmarkId.startsWith('prompt:')) {
    const cluster = getPromptClusters()[bookmarkId.slice(7)];
    if (!cluster) return;
    dom.clusterEditorTitle.textContent = 'Edit cluster';
    dom.clusterEditorCopy.textContent = 'Rename this cluster, change its color, or update the description.';
    dom.clusterEditorBookmarkId.value = bookmarkId;
    dom.clusterEditorName.value = cluster.label || bookmarkId.slice(7);
    dom.clusterEditorDescription.value = cluster.description || '';
    if (dom.clusterEditorColor) dom.clusterEditorColor.value = cluster.color || nextClusterColor(0);
  }
  dom.clusterEditorOverlay.classList.remove('hidden');
  setTimeout(() => (focusMode === 'description' ? dom.clusterEditorDescription : dom.clusterEditorName).focus(), 0);
}

function saveClusterEditorForm() {
  const bookmarkId = dom.clusterEditorBookmarkId.value;
  const name = dom.clusterEditorName.value.trim();
  const description = dom.clusterEditorDescription.value.trim();
  const color = dom.clusterEditorColor ? dom.clusterEditorColor.value : '';
  if (!name) return;
  if (!bookmarkId) {
    const clusters = getPromptClusters();
    const slug = createClusterSlug(name, clusters);
    clusters[slug] = {
      label: name,
      description,
      color: color || nextClusterColor(Object.keys(clusters).length),
      emailIds: [],
      overrides: {},
      createdAt: new Date().toISOString(),
      order: Object.keys(clusters).length
    };
    savePromptClusters(clusters);
    dom.clusterEditorOverlay.classList.add('hidden');
    renderApp();
    openBookmarkTab('prompt:' + slug);
    return;
  }
  if (bookmarkId.startsWith('prompt:')) {
    const clusters = getPromptClusters();
    const slug = bookmarkId.slice(7);
    if (!clusters[slug]) return;
    clusters[slug].label = name;
    clusters[slug].description = description;
    if (color) clusters[slug].color = color;
    savePromptClusters(clusters);
  }
  dom.clusterEditorOverlay.classList.add('hidden');
  renderApp();
}

function deleteBookmark(bookmarkId) {
  if (!bookmarkId || !confirm('Delete this cluster?')) return;
  if (bookmarkId.startsWith('prompt:')) {
    const clusters = getPromptClusters();
    delete clusters[bookmarkId.slice(7)];
    savePromptClusters(clusters);
  }
  const overrides = getBookmarkOverrides();
  Object.keys(overrides).forEach((emailId) => { if (overrides[emailId] === bookmarkId) delete overrides[emailId]; });
  saveBookmarkOverrides(overrides);
  appState.tabs = appState.tabs.filter((tab) => !(tab.type === 'clusterList' && tab.sourceType === 'bookmark' && tab.bookmarkId === bookmarkId));
  if (!appState.tabs.some((tab) => tab.id === appState.activeTabId)) appState.activeTabId = HOME_TAB_ID;
  renderApp();
}

async function openSettings() {
  if (typeof window.electronAPI === 'undefined') return;
  const config = await window.electronAPI.imap.getConfig();
  if (config) {
    dom.settingsForm.elements.host.value = config.host || '';
    dom.settingsForm.elements.port.value = config.port || 993;
    dom.settingsForm.elements.user.value = config.user || '';
    dom.settingsForm.elements.pass.value = config.pass || '';
  }
  dom.settingsStatus.textContent = '';
  dom.settingsOverlay.classList.remove('hidden');
}

async function testSettings() {
  if (typeof window.electronAPI === 'undefined') return;
  const config = {
    host: dom.settingsForm.elements.host.value.trim(),
    port: parseInt(dom.settingsForm.elements.port.value, 10) || 993,
    user: dom.settingsForm.elements.user.value.trim(),
    pass: dom.settingsForm.elements.pass.value
  };
  dom.settingsStatus.textContent = 'Testing…';
  const result = await window.electronAPI.imap.test(config);
  dom.settingsStatus.textContent = result.ok ? 'Connection OK.' : ('Failed: ' + (result.error || 'Unknown error'));
}

async function saveSettings(event) {
  event.preventDefault();
  if (typeof window.electronAPI === 'undefined') return;
  await window.electronAPI.imap.saveConfig({
    host: dom.settingsForm.elements.host.value.trim(),
    port: parseInt(dom.settingsForm.elements.port.value, 10) || 993,
    secure: true,
    user: dom.settingsForm.elements.user.value.trim(),
    pass: dom.settingsForm.elements.pass.value
  });
  dom.settingsStatus.textContent = 'Saved.';
}

function showTooltip(text, x, y) {
  if (!text) return;
  dom.tooltip.textContent = text;
  dom.tooltip.style.left = (x + 14) + 'px';
  dom.tooltip.style.top = (y + 14) + 'px';
  dom.tooltip.classList.remove('hidden');
}

function hideTooltip() { dom.tooltip.classList.add('hidden'); }

function openBookmarkContextMenu(bookmarkId, x, y) {
  contextBookmarkId = bookmarkId;
  dom.contextMenu.innerHTML = '<button type="button" class="context-menu-item" data-action="edit-bookmark-name">Rename</button><button type="button" class="context-menu-item" data-action="edit-bookmark-description">Edit description</button><button type="button" class="context-menu-item" data-action="delete-bookmark">Delete</button>';
  dom.contextMenu.style.left = x + 'px';
  dom.contextMenu.style.top = y + 'px';
  dom.contextMenu.classList.remove('hidden');
}

function closeBookmarkContextMenu() { dom.contextMenu.classList.add('hidden'); }

function handleDocumentInput(event) {
  const target = event.target;
  if (target.matches('[data-role="tab-query"]')) {
    const tab = appState.tabs.find((item) => item.id === target.dataset.tabId);
    if (!tab) return;
    tab.query = target.value;
    if (tab.type === 'clusterList') tab.visibleRows = LIST_INITIAL_ROWS;
    scheduleSearch(tab);
  } else if (target.matches('[data-role="compose-to"]')) {
    const tab = appState.tabs.find((item) => item.id === target.dataset.tabId && item.type === 'compose');
    if (!tab) return;
    tab.to = target.value;
    updateComposeTitleFromField(tab.id);
  } else if (target.matches('[data-role="compose-subject"]')) {
    const tab = appState.tabs.find((item) => item.id === target.dataset.tabId && item.type === 'compose');
    if (tab) tab.subject = target.value;
  } else if (target.matches('[data-role="compose-editor"]')) {
    updateComposeStateFromEditor(target.dataset.tabId, target);
  } else if (target === dom.thresholdSlider) {
    renderThresholdResults();
  }
}

function handleDocumentChange(event) {
  const target = event.target;
  if (target.matches('[data-role="row-select"]')) {
    updateClusterSelection(target.dataset.tabId, target.dataset.emailId, target.checked);
  } else if (target.matches('[data-role="row-move"]')) {
    if (target.value) moveEmailsToBookmark([target.dataset.emailId], target.value);
    target.value = '';
  } else if (target.matches('[data-role="batch-move"]')) {
    const tab = appState.tabs.find((item) => item.id === target.dataset.tabId && item.type === 'clusterList');
    if (tab && target.value) moveEmailsToBookmark(tab.selectedIds, target.value);
    target.value = '';
  }
}

function handleDocumentClick(event) {
  const actionEl = event.target.closest('[data-action]');
  if (!actionEl) {
    if (!event.target.closest('#bookmark-context-menu')) closeBookmarkContextMenu();
    return;
  }
  const action = actionEl.dataset.action;
  if (action !== 'edit-bookmark-name' && action !== 'edit-bookmark-description' && action !== 'delete-bookmark') closeBookmarkContextMenu();
  if (action === 'activate-tab') setActiveTab(actionEl.dataset.tabId);
  else if (action === 'close-tab') { event.stopPropagation(); closeTab(actionEl.dataset.tabId); }
  else if (action === 'open-bookmark-tab') openBookmarkTab(actionEl.dataset.bookmarkId);
  else if (action === 'open-compose-tab') openComposeTab();
  else if (action === 'focus-smart-cluster') focusSmartClusterPrompt();
  else if (action === 'open-settings') openSettings();
  else if (action === 'go-home') openHomeTab();
  else if (action === 'open-mailbox-tab') openSystemTab(actionEl.dataset.systemId, actionEl.dataset.label, actionEl.dataset.icon);
  else if (action === 'refresh-imap') refreshFromImap();
  else if (action === 'load-more') loadMoreImapEmails();
  else if (action === 'compute-embeddings') computeEmbeddings();
  else if (action === 'set-cluster-view') { const tab = appState.tabs.find((item) => item.id === actionEl.dataset.tabId && item.type === 'clusterList'); if (tab) { tab.viewMode = actionEl.dataset.viewMode; renderActiveView(); } }
  else if (action === 'open-email-tab') openEmailTab(actionEl.dataset.emailId || actionEl.closest('[data-email-id]')?.dataset.emailId);
  else if (action === 'clear-selection') clearClusterSelection(actionEl.dataset.tabId);
  else if (action === 'expand-cluster-list') expandClusterList(actionEl.dataset.tabId);
  else if (action === 'toggle-thread-message') { const tab = appState.tabs.find((item) => item.id === actionEl.dataset.tabId && item.type === 'emailThread'); if (tab) { tab.expandedMessageIds[actionEl.dataset.emailId] = !tab.expandedMessageIds[actionEl.dataset.emailId]; renderActiveView(); } }
  else if (action === 'switch-email-tab-message') { const tab = appState.tabs.find((item) => item.id === actionEl.dataset.tabId && item.type === 'emailThread'); if (tab) { tab.emailId = actionEl.dataset.emailId; tab.expandedMessageIds[actionEl.dataset.emailId] = true; const email = getEmailById(actionEl.dataset.emailId); if (email) { tab.title = truncate(email.subject || '(no subject)', 34); tab.iconName = 'mail-open'; } renderApp(); ensureEmailBodyLoaded(actionEl.dataset.emailId); } }
  else if (action === 'reply-email') replyEmail(actionEl.dataset.emailId);
  else if (action === 'forward-email') forwardEmail(actionEl.dataset.emailId);
  else if (action === 'delete-email') handleDeleteEmail(actionEl.dataset.emailId);
  else if (action === 'compose-command') execComposeCommand(actionEl.dataset.tabId, actionEl.dataset.command);
  else if (action === 'compose-link') insertComposeLink(actionEl.dataset.tabId);
  else if (action === 'compose-attach') attachFilesToCompose(actionEl.dataset.tabId);
  else if (action === 'remove-attachment') removeComposeAttachment(actionEl.dataset.tabId, actionEl.dataset.path);
  else if (action === 'send-compose') sendCompose(actionEl.dataset.tabId);
  else if (action === 'open-cluster-editor') openClusterEditor('', 'name');
  else if (action === 'close-overlay') document.getElementById(actionEl.dataset.overlayId).classList.add('hidden');
  else if (action === 'close-threshold-overlay') closeThresholdOverlay();
  else if (action === 'edit-bookmark-name') openClusterEditor(contextBookmarkId, 'name');
  else if (action === 'edit-bookmark-description') openClusterEditor(contextBookmarkId, 'description');
  else if (action === 'delete-bookmark') deleteBookmark(contextBookmarkId);
}

function handleContextMenu(event) {
  const bookmark = event.target.closest('.bookmark-pill, .bookmark-card');
  if (!bookmark || !bookmark.dataset.bookmarkId) return;
  event.preventDefault();
  openBookmarkContextMenu(bookmark.dataset.bookmarkId, event.clientX, event.clientY);
}

function handleMouseOver(event) {
  const tooltipTarget = event.target.closest('[data-tooltip]');
  if (!tooltipTarget) return;
  showTooltip(tooltipTarget.dataset.tooltip, event.clientX, event.clientY);
}

function handleMouseMove(event) {
  if (!dom.tooltip.classList.contains('hidden')) showTooltip(dom.tooltip.textContent, event.clientX, event.clientY);
}

function handleMouseOut(event) {
  if (!event.target.closest('[data-tooltip]')) return;
  hideTooltip();
}

function handleDocumentScroll(event) {
  const container = event.target && event.target.closest ? event.target.closest('.cluster-list-body[data-tab-id]') : null;
  if (!container) return;
  if ((container.scrollTop + container.clientHeight) < (container.scrollHeight - 240)) return;
  expandClusterList(container.dataset.tabId);
}

function focusSmartClusterPrompt() {
  const input = document.getElementById('prompt-cluster-input');
  if (!input) return;
  input.scrollIntoView({ behavior: 'smooth', block: 'center' });
  input.focus();
}

async function initialize() {
  renderApp();
  if (typeof window.electronAPI !== 'undefined' && window.electronAPI.embeddings && window.electronAPI.embeddings.onProgress) {
    window.electronAPI.embeddings.onProgress((progress) => setStatus(progress && progress.message ? progress.message : 'Working…', 'muted'));
  }
  if (typeof window.electronAPI !== 'undefined' && window.electronAPI.imap) {
    const config = await window.electronAPI.imap.getConfig();
    const accountKey = config && config.host && config.user ? config.host + '::' + config.user : '';
    if (config && config.host) imapHostLabel = config.host + (config.user ? ' / ' + config.user : '');
    if (accountKey) {
      imapEmails = await getCachedEmails(accountKey);
      invalidateEmailCollections();
      invalidateThreadCache();
      renderApp();
    }
    refreshFromImap();
  }
}

document.addEventListener('click', handleDocumentClick);
document.addEventListener('input', handleDocumentInput);
document.addEventListener('change', handleDocumentChange);
document.addEventListener('contextmenu', handleContextMenu);
document.addEventListener('mouseover', handleMouseOver);
document.addEventListener('mousemove', handleMouseMove);
document.addEventListener('mouseout', handleMouseOut);
document.addEventListener('scroll', handleDocumentScroll, true);
document.addEventListener('selectionchange', () => {
  const active = document.activeElement;
  if (active && active.matches && active.matches('[data-role="compose-editor"]')) saveComposeSelection(active.dataset.tabId);
});
document.addEventListener('mousedown', (event) => {
  const tool = event.target.closest('.compose-tool-btn');
  if (tool) event.preventDefault();
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    closeBookmarkContextMenu();
    hideTooltip();
    dom.settingsOverlay.classList.add('hidden');
    dom.clusterEditorOverlay.classList.add('hidden');
    closeThresholdOverlay();
  }
});

dom.thresholdResults.addEventListener('click', (event) => {
  const row = event.target.closest('[data-role="threshold-row"]');
  if (!row || !pendingPromptCluster) return;
  const id = row.dataset.emailId;
  const threshold = parseFloat(dom.thresholdSlider.value || '0.3');
  const item = pendingPromptCluster.scored.find((entry) => entry.id === id);
  if (!item) return;
  const included = pendingPromptCluster.overrides[id] === true || (pendingPromptCluster.overrides[id] !== false && item.sim >= threshold);
  pendingPromptCluster.overrides[id] = included ? false : true;
  renderThresholdResults();
});
dom.thresholdCreate.addEventListener('click', commitThresholdCluster);
dom.settingsForm.addEventListener('submit', saveSettings);
dom.clusterEditorForm.addEventListener('submit', (event) => { event.preventDefault(); saveClusterEditorForm(); });
document.getElementById('settings-test').addEventListener('click', testSettings);
document.addEventListener('submit', (event) => {
  const form = event.target;
  if (form && form.id === 'prompt-cluster-form') {
    event.preventDefault();
    createPromptClusterFromPrompt(document.getElementById('prompt-cluster-input').value);
  }
});

initialize();
