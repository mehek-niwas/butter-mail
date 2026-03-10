const STORAGE_KEY = 'butter-mail-emails';
const EMBEDDINGS_KEY = 'butter-mail-embeddings';
const CATEGORIES_KEY = 'butter-mail-categories';
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
  clusterEditorDescription: document.getElementById('cluster-editor-description'),
  thresholdOverlay: document.getElementById('cluster-threshold-overlay'),
  thresholdTitle: document.getElementById('cluster-threshold-title'),
  thresholdResults: document.getElementById('cluster-threshold-results'),
  thresholdSlider: document.getElementById('cluster-threshold-slider'),
  thresholdValue: document.getElementById('cluster-threshold-value'),
  thresholdCount: document.getElementById('cluster-threshold-count'),
  thresholdCreate: document.getElementById('cluster-threshold-create'),
  themeToggle: document.getElementById('theme-toggle-btn')
};

const mailboxShortcuts = [
  { id: 'home', label: 'Home', icon: 'H' },
  { id: 'all-mail-page', label: 'All', icon: 'A' },
  { id: 'mailbox:INBOX', label: 'Inbox', icon: 'I' },
  { id: 'mailbox:Sent', label: 'Sent', icon: 'S' },
  { id: 'mailbox:Drafts', label: 'Drafts', icon: 'D' },
  { id: 'mailbox:Trash', label: 'Trash', icon: 'T' }
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
let storedEmails = getJson(STORAGE_KEY, []);
let storedEmbeddings = getJson(EMBEDDINGS_KEY, {});
let storedCategories = getJson(CATEGORIES_KEY, { assignments: {}, meta: {} });
let storedPromptClusters = getJson(PROMPT_CLUSTERS_KEY, {});
let storedBookmarkOverrides = getJson(BOOKMARK_OVERRIDES_KEY, {});
let allEmailsCache = null;
let allEmailsByIdCache = null;
let bookmarkDefinitionsCache = null;
let systemEmailCache = {};
let emailCollectionRevision = 0;
let bookmarkRevision = 0;

const appState = {
  theme: localStorage.getItem(THEME_KEY) === 'light' ? 'light' : 'dark',
  tabs: [createHomeTab()],
  activeTabId: HOME_TAB_ID,
  status: '',
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
function getCategories() { return storedCategories; }
function saveCategories(cats) {
  storedCategories = cats || { assignments: {}, meta: {} };
  saveJson(CATEGORIES_KEY, storedCategories);
  invalidateBookmarkDefinitions();
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
  return { id: HOME_TAB_ID, type: 'home', title: 'Home', iconLabel: 'H', closable: false, query: '', searchResults: null, searchLoading: false };
}

function createClusterTab(options) {
  return {
    id: makeId('tab-cluster'),
    type: 'clusterList',
    sourceType: options.sourceType,
    bookmarkId: options.bookmarkId || '',
    systemId: options.systemId || '',
    title: options.title,
    iconLabel: options.iconLabel || initialsFromText(options.title),
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
    iconLabel: initialsFromText(email && (email.fromEmail || email.from || email.subject) ? (email.fromEmail || email.from || email.subject) : 'E'),
    closable: true,
    expandedMessageIds: { [emailId]: true }
  };
}

function createComposeTab(seed) {
  const tab = {
    id: makeId('tab-compose'),
    type: 'compose',
    title: 'New Email',
    iconLabel: 'P',
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

function initialsFromText(value) {
  const parts = String(value || '').trim().replace(/<[^>]+>/g, '').split(/[\s@._-]+/).filter(Boolean);
  if (!parts.length) return '?';
  if (parts.length === 1) return parts[0].slice(0, 1).toUpperCase();
  return (parts[0].slice(0, 1) + parts[1].slice(0, 1)).toUpperCase();
}

function hashString(value) {
  let hash = 0;
  const text = String(value || '');
  for (let index = 0; index < text.length; index += 1) {
    hash = ((hash << 5) - hash) + text.charCodeAt(index);
    hash |= 0;
  }
  return Math.abs(hash);
}

function iconMarkup(seed, label) {
  const variant = (hashString(seed || label) % 5) + 1;
  return '<span class="icon-chip variant-' + variant + '">' + escapeHtml(label) + '</span>';
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

function openSystemTab(systemId, label, icon) {
  return addTab(createClusterTab({ sourceType: 'system', systemId, title: label, iconLabel: icon || initialsFromText(label) }), 'system:' + systemId);
}

function openAllMailTab() { return openSystemTab('all-mail', 'All', 'A'); }

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
  dom.body.setAttribute('data-theme', appState.theme);
  if (dom.themeToggle) dom.themeToggle.textContent = appState.theme === 'dark' ? 'Light' : 'Dark';
  if (next.tabs !== false) renderTabStrip();
  if (next.bookmarks !== false) renderBookmarkBar();
  if (next.view !== false) renderActiveView();
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

function sortedAutoCategories() {
  const cats = getCategories();
  return Object.keys(cats.meta || {}).filter((key) => key !== 'noise').sort((left, right) => {
    const metaA = cats.meta[left] || {};
    const metaB = cats.meta[right] || {};
    const orderA = typeof metaA.order === 'number' ? metaA.order : Number.MAX_SAFE_INTEGER;
    const orderB = typeof metaB.order === 'number' ? metaB.order : Number.MAX_SAFE_INTEGER;
    if (orderA !== orderB) return orderA - orderB;
    return String(metaA.name || left).localeCompare(String(metaB.name || right));
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
  if (bookmarkId.startsWith('auto:')) {
    return getCategories().assignments[email.id] === bookmarkId.slice(5);
  }
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
  const cats = getCategories();
  sortedPromptClusters().forEach(([slug, cluster]) => {
    const memberIds = getPromptClusterMemberIds(cluster);
    bookmarks.push({
      id: 'prompt:' + slug,
      slug,
      label: cluster.label || slug,
      description: cluster.description || '',
      kind: 'user',
      count: countEmailsForBookmark('prompt:' + slug, memberIds)
    });
  });
  sortedAutoCategories().forEach((catId) => {
    const meta = cats.meta[catId] || {};
    bookmarks.push({
      id: 'auto:' + catId,
      slug: catId,
      label: meta.name || catId,
      description: meta.description || '',
      kind: 'auto',
      count: countEmailsForBookmark('auto:' + catId)
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
    iconLabel: initialsFromText(bookmark.label)
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
  const key = tab.systemId || 'all-mail';
  if (systemEmailCache[key]) return systemEmailCache[key];
  const emails = getAllEmails();
  let filtered = emails;
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
  dom.tabStrip.innerHTML = appState.tabs.map((tab) => {
    const active = tab.id === appState.activeTabId ? ' active' : '';
    return '<div class="browser-tab' + active + '" data-action="activate-tab" data-tab-id="' + escapeHtml(tab.id) + '" role="button" tabindex="0">' +
      '<span class="tab-favicon">' + iconMarkup(tab.title || tab.iconLabel, tab.iconLabel || initialsFromText(tab.title || 'T')) + '</span>' +
      '<span class="browser-tab-title">' + escapeHtml(tab.title) + '</span>' +
      (tab.closable ? '<button type="button" class="tab-close" data-action="close-tab" data-tab-id="' + escapeHtml(tab.id) + '" aria-label="Close tab">×</button>' : '') +
    '</div>';
  }).join('');
}

function renderBookmarkBar() {
  const active = getActiveTab();
  const currentBookmarkId = active && active.type === 'clusterList' && active.sourceType === 'bookmark' ? active.bookmarkId : '';
  dom.bookmarkBar.innerHTML = getBookmarkDefinitions().map((bookmark) => {
    const tooltip = (bookmark.kind === 'auto' ? 'Auto-created cluster' : 'User-created cluster') + (bookmark.description ? '\n' + bookmark.description : '');
    const activeClass = bookmark.id === currentBookmarkId ? ' active' : '';
    return '<button type="button" class="bookmark-pill' + activeClass + '" data-action="open-bookmark-tab" data-bookmark-id="' + escapeHtml(bookmark.id) + '" data-tooltip="' + escapeHtml(tooltip) + '">' +
      '<span class="bookmark-favicon">' + iconMarkup(bookmark.label, initialsFromText(bookmark.label)) + '</span>' +
      '<span>' + escapeHtml(bookmark.label) + '</span>' +
    '</button>';
  }).join('');
}

function renderBookmarkCard(bookmark) {
  return '<button type="button" class="bookmark-card" data-action="open-bookmark-tab" data-bookmark-id="' + escapeHtml(bookmark.id) + '" data-tooltip="' + escapeHtml((bookmark.kind === 'auto' ? 'Auto-created cluster' : 'User-created cluster') + (bookmark.description ? '\n' + bookmark.description : '')) + '">' +
    '<div class="bookmark-card-header">' + iconMarkup(bookmark.label, initialsFromText(bookmark.label)) + '<span class="bookmark-card-count">' + String(bookmark.count) + ' email(s)</span></div>' +
    '<div><h3 class="bookmark-card-name">' + escapeHtml(bookmark.label) + '</h3><p class="bookmark-card-description">' + escapeHtml(bookmark.description || (bookmark.kind === 'auto' ? 'Auto-created from clustering.' : 'User-created bookmark cluster.')) + '</p></div>' +
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
  const allEmails = sortEmails(getAllEmails(), false).slice(0, 12);
  return '<section class="tab-view home-view">' +
    '<div class="home-hero">' +
      '<div class="hero-logo-row"><img src="assets/butter-mail-logo.webp" alt="Butter Mail" class="hero-logo" /><img src="assets/bread-butter.webp" alt="" class="hero-logo hero-logo-secondary" aria-hidden="true" /></div>' +
      '<h1 class="hero-title">Welcome to your spread</h1>' +
      '<p class="hero-copy">Browse everything at once, or move through clusters and threads without the usual inbox overload.</p>' +
      '<div class="hero-tagline">the better mail, with your spread front and center</div>' +
    '</div>' +
    '<div class="home-search-row">' +
      '<button type="button" class="primary-btn" data-action="open-compose-tab">Compose</button>' +
      '<div class="home-search-shell"><span class="meta-pill">Search</span><input type="text" class="search-field" data-role="tab-query" data-tab-id="' + escapeHtml(tab.id) + '" placeholder="Search across your spread" value="' + escapeHtml(tab.query || '') + '" /></div>' +
      '<button type="button" class="secondary-btn" data-action="open-all-mail-tab">All</button>' +
    '</div>' +
    '<div class="home-shortcuts">' +
      mailboxShortcuts.map((shortcut) => {
        if (shortcut.id === 'home') return '<button type="button" class="home-shortcut" data-action="go-home">' + iconMarkup(shortcut.label, shortcut.icon) + '<span>' + escapeHtml(shortcut.label) + '</span></button>';
        if (shortcut.id === 'all-mail-page') return '<button type="button" class="home-shortcut" data-action="open-all-mail-tab">' + iconMarkup(shortcut.label, shortcut.icon) + '<span>' + escapeHtml(shortcut.label) + '</span></button>';
        return '<button type="button" class="home-shortcut" data-action="open-mailbox-tab" data-system-id="' + escapeHtml(shortcut.id) + '" data-label="' + escapeHtml(shortcut.label) + '" data-icon="' + escapeHtml(shortcut.icon) + '">' + iconMarkup(shortcut.label, shortcut.icon) + '<span>' + escapeHtml(shortcut.label) + '</span></button>';
      }).join('') +
    '</div>' +
    (tab.query ? '<section class="surface-card"><h2 class="surface-title">Search results</h2><p class="surface-copy">' + (tab.searchLoading ? 'Searching your spread…' : String(searchResults.length) + ' result(s)') + '</p>' + renderSearchPreviewList(searchResults) + '</section>' : '') +
    '<div class="home-grid">' +
      '<section class="surface-card"><h2 class="surface-title">Bookmarks</h2><p class="surface-copy">' + String(bookmarks.length) + ' cluster(s) available.</p><div class="bookmark-grid">' + bookmarks.map(renderBookmarkCard).join('') + '</div></section>' +
      '<aside class="automation-stack">' +
        '<section class="surface-card"><h2 class="surface-title">Short actions</h2><p class="surface-copy">Keep the main flow calm. Use the secondary panel for heavier actions.</p><div class="toolbar-row"><button type="button" class="quick-action" data-action="refresh-imap">Refresh</button><button type="button" class="quick-action" data-action="compute-embeddings">Compute embeddings</button><button type="button" class="quick-action" data-action="recluster">Re-cluster</button></div><p class="status-copy ' + escapeHtml(appState.statusTone) + '" id="global-status">' + escapeHtml(appState.status || '') + '</p></section>' +
        '<section class="surface-card"><h2 class="surface-title">Create a smart cluster</h2><p class="surface-copy">Turn a short concept into a bookmark backed by semantic similarity.</p><form class="prompt-cluster-form" id="prompt-cluster-form"><input type="text" class="compose-field" id="prompt-cluster-input" placeholder="e.g. invoices, job hunt, design reviews" /><button type="submit" class="primary-btn">Create</button></form></section>' +
      '</aside>' +
    '</div>' +
    '<section class="surface-card all-mail-surface"><div class="all-mail-header"><div><h2 class="surface-title">All emails</h2><p class="surface-copy">A clean recent view from every mailbox, directly inside the Home tab.</p></div><button type="button" class="secondary-btn" data-action="open-all-mail-tab">Open All tab</button></div>' + renderClusterRows({ id: HOME_TAB_ID, selectedIds: [], type: 'home' }, allEmails, { selectable: false }) + '</section>' +
  '</section>';
}

function renderBookmarkOptions(includeBase) {
  const bookmarks = getBookmarkDefinitions();
  return '<option value="">Move to…</option>' +
    (includeBase ? '<option value="' + BASE_OVERRIDE_VALUE + '">Base classification</option>' : '') +
    bookmarks.map((bookmark) => '<option value="' + escapeHtml(bookmark.id) + '">' + escapeHtml(bookmark.label) + '</option>').join('');
}

function renderBatchToolbar(tab) {
  return '<div class="cluster-batch-toolbar"><span class="toolbar-note">' + String(tab.selectedIds.length) + ' selected</span><select data-role="batch-move" data-tab-id="' + escapeHtml(tab.id) + '">' + renderBookmarkOptions(true) + '</select><button type="button" class="secondary-btn" data-action="clear-selection" data-tab-id="' + escapeHtml(tab.id) + '">Clear</button></div>';
}

function renderGraphShell(emails) {
  if (!Object.keys(getEmbeddings()).length || !Object.keys(pcaPoints || {}).length) return '<div class="state-block">Compute embeddings first to use graph mode.</div>';
  if (!emails.length) return '<div class="state-block">Nothing to graph in this tab.</div>';
  return '<div class="graph-shell"><div class="graph-container" id="graph-container"><canvas id="graph-canvas"></canvas><div class="graph-axis-legend"><span class="graph-axis-label graph-axis-x">X</span><span class="graph-axis-label graph-axis-y">Y</span><span class="graph-axis-label graph-axis-z">Z</span></div><div class="graph-coords-legend" id="graph-coords">X: — Y: — Z: —</div><div class="graph-tooltip hidden" id="graph-tooltip"></div></div></div>';
}

function renderClusterRows(tab, emails, options) {
  if (!emails.length) return '<div class="state-block">No emails here yet.</div>';
  ensureThreadCacheBuilt();
  const selectable = !options || options.selectable !== false;
  const renderedCount = getRenderedEmailCount(tab, emails.length);
  const visibleEmails = emails.slice(0, renderedCount);
  const remainingCount = Math.max(0, emails.length - renderedCount);
  return '<div class="cluster-list-body"' + (tab.type === 'clusterList' ? ' data-tab-id="' + escapeHtml(tab.id) + '"' : '') + '>' + visibleEmails.map((email) => {
    const selected = tab.selectedIds.includes(email.id) ? ' checked' : '';
    const threadSize = threadSizesByEmailId[email.id] || 1;
    const sender = email.from || email.fromEmail || '(unknown sender)';
    const mailbox = normalizeMailbox(email);
    return '<div class="cluster-row" data-email-id="' + escapeHtml(email.id) + '">' +
      (selectable ? '<input class="cluster-row-checkbox" type="checkbox" data-role="row-select" data-tab-id="' + escapeHtml(tab.id) + '" data-email-id="' + escapeHtml(email.id) + '"' + selected + ' />' : '') +
      '<div class="row-main" data-action="open-email-tab" data-email-id="' + escapeHtml(email.id) + '">' +
        '<div class="row-meta-line">' +
          '<span class="cluster-row-leading">' + iconMarkup(email.fromEmail || sender || email.subject || 'E', initialsFromText(email.fromEmail || sender || 'E')) + '</span>' +
          '<span class="row-sender">' + escapeHtml(sender) + '</span>' +
          '<span class="row-separator">›</span>' +
          '<span class="row-mailbox">' + escapeHtml(mailbox || 'mail') + '</span>' +
          (threadSize > 1 ? '<span class="thread-indicator">thread · ' + String(threadSize) + '</span>' : '') +
          '<span class="row-date">' + escapeHtml(formatDate(email.date)) + '</span>' +
        '</div>' +
        '<div class="row-copy">' +
          '<div class="row-subject">' + escapeHtml(email.subject || '(no subject)') + '</div>' +
          '<div class="row-preview">' + escapeHtml(getEmailPreview(email)) + '</div>' +
        '</div>' +
      '</div>' +
      (selectable ? '<div class="cluster-row-actions"><select class="cluster-row-move" data-role="row-move" data-email-id="' + escapeHtml(email.id) + '">' + renderBookmarkOptions(true) + '</select><button type="button" class="cluster-row-open" data-action="open-email-tab" data-email-id="' + escapeHtml(email.id) + '">Open</button></div>' : '<div class="cluster-row-actions cluster-row-actions-static"><button type="button" class="cluster-row-open" data-action="open-email-tab" data-email-id="' + escapeHtml(email.id) + '">Open</button></div>') +
    '</div>';
  }).join('') + (remainingCount ? '<button type="button" class="cluster-list-more" data-action="expand-cluster-list" data-tab-id="' + escapeHtml(tab.id) + '">Show ' + String(Math.min(LIST_ROW_STEP, remainingCount)) + ' more emails (' + String(remainingCount) + ' remaining)</button>' : '') + '</div>';
}

function renderClusterListView(tab) {
  const emails = getEmailsForTab(tab);
  const bookmark = tab.sourceType === 'bookmark' ? getBookmarkById(tab.bookmarkId) : null;
  const title = bookmark ? bookmark.label : tab.title;
  const subtitle = bookmark ? ((bookmark.description || (bookmark.kind === 'auto' ? 'Auto-created cluster.' : 'User cluster.')) + ' ' + String(emails.length) + ' email(s).') : (tab.systemId === 'all-mail' ? 'Every email in one search-style list view.' : String(emails.length) + ' email(s).');
  const canLoadMore = window.electronAPI && (tab.systemId === 'all-mail' || tab.systemId === 'mailbox:INBOX');
  const renderedCount = tab.viewMode === 'graph' ? emails.length : getRenderedEmailCount(tab, emails.length);
  return '<section class="tab-view cluster-view">' +
    '<div class="view-header"><div><h1 class="view-title">' + escapeHtml(title) + '</h1><p class="view-subtitle">' + escapeHtml(subtitle) + '</p></div><div class="segmented"><button type="button" class="segmented-btn' + (tab.viewMode === 'list' ? ' active' : '') + '" data-action="set-cluster-view" data-tab-id="' + escapeHtml(tab.id) + '" data-view-mode="list">List</button><button type="button" class="segmented-btn' + (tab.viewMode === 'graph' ? ' active' : '') + '" data-action="set-cluster-view" data-tab-id="' + escapeHtml(tab.id) + '" data-view-mode="graph">Graph</button></div></div>' +
    '<div class="cluster-list-panel"><div class="cluster-toolbar"><div class="cluster-search-shell"><span class="meta-pill">Search</span><input type="text" class="search-field" data-role="tab-query" data-tab-id="' + escapeHtml(tab.id) + '" placeholder="Filter this tab" value="' + escapeHtml(tab.query || '') + '" /></div><div class="toolbar-row">' + (canLoadMore ? '<button type="button" class="toolbar-pill" data-action="load-more">Load more</button>' : '') + '<button type="button" class="toolbar-pill" data-action="refresh-imap">Refresh</button></div></div>' + (tab.selectedIds.length ? renderBatchToolbar(tab) : '') + (tab.searchLoading ? '<p class="toolbar-note">Searching…</p>' : '') + (tab.viewMode === 'graph' ? renderGraphShell(emails) : '<p class="toolbar-note">Showing ' + String(renderedCount) + ' of ' + String(emails.length) + ' emails.</p>' + renderClusterRows(tab, emails)) + '</div>' +
  '</section>';
}

function renderThreadMessageCard(tab, message, active) {
  const expanded = tab.expandedMessageIds[message.id] || active;
  return '<article class="message-card"><div class="message-card-head" data-action="toggle-thread-message" data-tab-id="' + escapeHtml(tab.id) + '" data-email-id="' + escapeHtml(message.id) + '"><div><div class="row-subject">' + escapeHtml(message.subject || '(no subject)') + '</div><div class="row-sender">' + escapeHtml(message.from || message.fromEmail || '') + '</div></div><div class="row-date">' + escapeHtml(formatDateTime(message.date)) + '</div></div>' + (expanded ? '<div class="message-card-body"><div class="message-body">' + renderEmailBody(message.body || '', !!message.bodyIsHtml) + '</div></div>' : '') + '</article>';
}

function renderEmailView(tab) {
  const email = getEmailById(tab.emailId);
  if (!email) return '<section class="tab-view"><div class="state-block">This email is no longer available.</div></section>';
  ensureThreadCacheBuilt();
  const thread = threadIndexByEmailId[email.id] || [email];
  return '<section class="tab-view message-view"><div class="message-shell"><div class="message-reader"><h1 class="message-subject">' + escapeHtml(email.subject || '(no subject)') + '</h1><div class="message-meta-grid"><div class="message-meta-card"><span class="meta-label">From</span>' + escapeHtml(email.from || email.fromEmail || '') + '</div><div class="message-meta-card"><span class="meta-label">To</span>' + escapeHtml(email.toDisplay || email.to || '') + '</div><div class="message-meta-card"><span class="meta-label">Date</span>' + escapeHtml(formatDateTime(email.date)) + '</div></div><div class="message-actions"><button type="button" class="message-action-btn" data-action="reply-email" data-email-id="' + escapeHtml(email.id) + '">Reply</button><button type="button" class="message-action-btn" data-action="forward-email" data-email-id="' + escapeHtml(email.id) + '">Forward</button><select class="cluster-row-move" data-role="row-move" data-email-id="' + escapeHtml(email.id) + '">' + renderBookmarkOptions(true) + '</select><button type="button" class="message-action-btn" data-action="delete-email" data-email-id="' + escapeHtml(email.id) + '">Delete</button></div><div class="message-thread">' + thread.map((message) => renderThreadMessageCard(tab, message, message.id === email.id)).join('') + '</div></div><aside class="message-sidebar"><h2 class="surface-title">Thread</h2><p class="surface-copy">' + String(thread.length) + ' message(s) in this conversation.</p><div class="search-preview-list">' + thread.map((message) => '<button type="button" class="search-preview-row" data-action="switch-email-tab-message" data-tab-id="' + escapeHtml(tab.id) + '" data-email-id="' + escapeHtml(message.id) + '"><div class="row-subject">' + escapeHtml(message.subject || '(no subject)') + '</div><div class="row-sender">' + escapeHtml(message.from || message.fromEmail || '') + '</div><div class="row-date">' + escapeHtml(formatDate(message.date)) + '</div></button>').join('') + '</div></aside></div></section>';
}

function renderAttachmentList(tab) {
  if (!Array.isArray(tab.attachments) || !tab.attachments.length) return '<div class="state-block">No attachments yet.</div>';
  return tab.attachments.map((attachment) => {
    const parts = String(attachment.path || '').split(/[/\\]/);
    const label = attachment.filename || parts[parts.length - 1] || attachment.path || 'attachment';
    return '<div class="compose-attachment-item"><span>' + escapeHtml(label) + '</span><button type="button" class="secondary-btn" data-action="remove-attachment" data-tab-id="' + escapeHtml(tab.id) + '" data-path="' + escapeHtml(attachment.path) + '">Remove</button></div>';
  }).join('');
}

function renderComposeView(tab) {
  return '<section class="tab-view compose-view"><div class="view-header"><div><h1 class="view-title">' + escapeHtml(tab.title) + '</h1><p class="view-subtitle">Write in a dedicated tab, then switch back whenever you need.</p></div></div><div class="compose-shell"><div class="compose-panel"><div class="compose-fields"><input type="text" class="compose-field" data-role="compose-to" data-tab-id="' + escapeHtml(tab.id) + '" placeholder="To" value="' + escapeHtml(tab.to) + '" /><input type="text" class="compose-field" data-role="compose-subject" data-tab-id="' + escapeHtml(tab.id) + '" placeholder="Subject" value="' + escapeHtml(tab.subject) + '" /></div><div class="compose-toolbar"><button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="bold"><strong>B</strong></button><button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="italic"><em>I</em></button><button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="underline"><u>U</u></button><button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="insertUnorderedList">• List</button><button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="insertOrderedList">1. List</button><button type="button" class="compose-tool-btn" data-action="compose-link" data-tab-id="' + escapeHtml(tab.id) + '">Link</button><button type="button" class="compose-tool-btn" data-action="compose-command" data-tab-id="' + escapeHtml(tab.id) + '" data-command="unlink">Unlink</button></div><div class="compose-editor" contenteditable="true" spellcheck="true" data-role="compose-editor" data-tab-id="' + escapeHtml(tab.id) + '" data-placeholder="Write your message...">' + sanitizeComposeHtml(tab.bodyHtml) + '</div><div class="compose-actions"><button type="button" class="secondary-btn" data-action="compose-attach" data-tab-id="' + escapeHtml(tab.id) + '">Attach files</button><button type="button" class="primary-btn" data-action="send-compose" data-tab-id="' + escapeHtml(tab.id) + '">' + (tab.sending ? 'Sending…' : 'Send') + '</button></div><p class="compose-status ' + escapeHtml(tab.statusTone || 'muted') + '">' + escapeHtml(tab.status || '') + '</p></div><aside class="compose-side"><h2 class="surface-title">Attachments</h2><p class="surface-copy">Minimal formatting, tight spacing, and a dedicated compose tab.</p><div class="compose-attachment-list">' + renderAttachmentList(tab) + '</div></aside></div></section>';
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
      emailsById[email.id] = { ...email, categoryId: getCategories().assignments[email.id] || null };
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

window.getCategoryColor = function (catId) {
  return (getCategories().meta[catId] || {}).color || '#7c83ff';
};

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

function buildThreadRepsForClustering(allEmails, embeddings) {
  const emails = allEmails.filter((email) => embeddings[email.id]);
  if (!emails.length) return { repIds: [], repToMembers: {} };
  const threads = window.ThreadView && typeof window.ThreadView.buildThreads === 'function' ? window.ThreadView.buildThreads(emails) : emails.map((email) => [email]);
  const repIds = [];
  const repToMembers = {};
  threads.forEach((thread) => {
    if (!thread || !thread.length) return;
    const members = thread.filter((email) => embeddings[email.id]);
    if (!members.length) return;
    const rep = members[members.length - 1];
    repIds.push(rep.id);
    repToMembers[rep.id] = members.map((email) => email.id);
  });
  return { repIds, repToMembers };
}

function expandClusterAssignmentsToThreads(clusterRes, repToMembers, embeddings) {
  const assignments = {};
  const meta = { ...(clusterRes.meta || {}) };
  Object.keys(repToMembers).forEach((repId) => {
    const catId = (clusterRes.assignments || {})[repId] || 'noise';
    repToMembers[repId].forEach((emailId) => { assignments[emailId] = catId; });
  });
  Object.keys(embeddings || {}).forEach((emailId) => { if (!assignments[emailId]) assignments[emailId] = 'noise'; });
  if (!meta.noise) meta.noise = { name: 'Uncategorized', color: '#909090' };
  return { assignments, meta };
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
  await ensureThreadHeaders();
  const threadData = buildThreadRepsForClustering(getAllEmails(), embeddingResult.embeddings || {});
  const clusterResult = await window.electronAPI.embeddings.cluster(embeddingResult.embeddings || {}, threadData.repIds);
  if (!clusterResult.ok) {
    alert('Clustering failed: ' + (clusterResult.error || 'Unknown error'));
    return;
  }
  saveCategories(expandClusterAssignmentsToThreads(clusterResult, threadData.repToMembers, embeddingResult.embeddings || {}));
  setStatus('Embeddings and clusters updated.', 'success');
  renderApp();
}

async function recluster() {
  if (typeof window.electronAPI === 'undefined') {
    alert('Re-cluster requires the Electron app.');
    return;
  }
  const embeddings = getEmbeddings();
  if (!Object.keys(embeddings).length) {
    alert('Compute embeddings first.');
    return;
  }
  setStatus('Re-clustering…', 'muted');
  await ensureThreadHeaders();
  const threadData = buildThreadRepsForClustering(getAllEmails(), embeddings);
  const clusterResult = await window.electronAPI.embeddings.cluster(embeddings, threadData.repIds);
  if (!clusterResult.ok) {
    alert('Re-cluster failed: ' + (clusterResult.error || 'Unknown error'));
    return;
  }
  saveCategories(expandClusterAssignmentsToThreads(clusterResult, threadData.repToMembers, embeddings));
  setStatus('Clusters refreshed.', 'success');
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
    dom.clusterEditorCopy.textContent = 'Create a user-managed bookmark for email triage.';
    dom.clusterEditorBookmarkId.value = '';
    dom.clusterEditorName.value = '';
    dom.clusterEditorDescription.value = '';
  } else if (bookmarkId.startsWith('prompt:')) {
    const cluster = getPromptClusters()[bookmarkId.slice(7)];
    if (!cluster) return;
    dom.clusterEditorTitle.textContent = 'Edit cluster';
    dom.clusterEditorCopy.textContent = 'Rename this cluster or update its tooltip description.';
    dom.clusterEditorBookmarkId.value = bookmarkId;
    dom.clusterEditorName.value = cluster.label || bookmarkId.slice(7);
    dom.clusterEditorDescription.value = cluster.description || '';
  } else if (bookmarkId.startsWith('auto:')) {
    const meta = (getCategories().meta || {})[bookmarkId.slice(5)];
    if (!meta) return;
    dom.clusterEditorTitle.textContent = 'Edit auto cluster';
    dom.clusterEditorCopy.textContent = 'Rename the clustered bookmark or add a description.';
    dom.clusterEditorBookmarkId.value = bookmarkId;
    dom.clusterEditorName.value = meta.name || bookmarkId.slice(5);
    dom.clusterEditorDescription.value = meta.description || '';
  }
  dom.clusterEditorOverlay.classList.remove('hidden');
  setTimeout(() => (focusMode === 'description' ? dom.clusterEditorDescription : dom.clusterEditorName).focus(), 0);
}

function saveClusterEditorForm() {
  const bookmarkId = dom.clusterEditorBookmarkId.value;
  const name = dom.clusterEditorName.value.trim();
  const description = dom.clusterEditorDescription.value.trim();
  if (!name) return;
  if (!bookmarkId) {
    const clusters = getPromptClusters();
    const slug = createClusterSlug(name, clusters);
    clusters[slug] = { label: name, description, emailIds: [], overrides: {}, createdAt: new Date().toISOString(), order: Object.keys(clusters).length };
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
    savePromptClusters(clusters);
  } else if (bookmarkId.startsWith('auto:')) {
    const cats = getCategories();
    const catId = bookmarkId.slice(5);
    if (!cats.meta[catId]) return;
    cats.meta[catId].name = name;
    cats.meta[catId].description = description;
    saveCategories(cats);
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
  } else if (bookmarkId.startsWith('auto:')) {
    const cats = getCategories();
    const catId = bookmarkId.slice(5);
    delete cats.meta[catId];
    Object.keys(cats.assignments || {}).forEach((emailId) => { if (cats.assignments[emailId] === catId) delete cats.assignments[emailId]; });
    saveCategories(cats);
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
  else if (action === 'toggle-theme') { appState.theme = appState.theme === 'dark' ? 'light' : 'dark'; localStorage.setItem(THEME_KEY, appState.theme); renderApp(); }
  else if (action === 'open-settings') openSettings();
  else if (action === 'go-home') openHomeTab();
  else if (action === 'open-all-mail-tab') openAllMailTab();
  else if (action === 'open-mailbox-tab') openSystemTab(actionEl.dataset.systemId, actionEl.dataset.label, actionEl.dataset.icon);
  else if (action === 'refresh-imap') refreshFromImap();
  else if (action === 'load-more') loadMoreImapEmails();
  else if (action === 'compute-embeddings') computeEmbeddings();
  else if (action === 'recluster') recluster();
  else if (action === 'set-cluster-view') { const tab = appState.tabs.find((item) => item.id === actionEl.dataset.tabId && item.type === 'clusterList'); if (tab) { tab.viewMode = actionEl.dataset.viewMode; renderActiveView(); } }
  else if (action === 'open-email-tab') openEmailTab(actionEl.dataset.emailId || actionEl.closest('[data-email-id]')?.dataset.emailId);
  else if (action === 'clear-selection') clearClusterSelection(actionEl.dataset.tabId);
  else if (action === 'expand-cluster-list') expandClusterList(actionEl.dataset.tabId);
  else if (action === 'toggle-thread-message') { const tab = appState.tabs.find((item) => item.id === actionEl.dataset.tabId && item.type === 'emailThread'); if (tab) { tab.expandedMessageIds[actionEl.dataset.emailId] = !tab.expandedMessageIds[actionEl.dataset.emailId]; renderActiveView(); } }
  else if (action === 'switch-email-tab-message') { const tab = appState.tabs.find((item) => item.id === actionEl.dataset.tabId && item.type === 'emailThread'); if (tab) { tab.emailId = actionEl.dataset.emailId; tab.expandedMessageIds[actionEl.dataset.emailId] = true; const email = getEmailById(actionEl.dataset.emailId); if (email) { tab.title = truncate(email.subject || '(no subject)', 34); tab.iconLabel = initialsFromText(email.fromEmail || email.from || email.subject || 'E'); } renderApp(); ensureEmailBodyLoaded(actionEl.dataset.emailId); } }
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

async function initialize() {
  renderApp();
  if (typeof window.electronAPI !== 'undefined' && window.electronAPI.embeddings && window.electronAPI.embeddings.onProgress) {
    window.electronAPI.embeddings.onProgress((progress) => setStatus(progress && progress.message ? progress.message : 'Working…', 'muted'));
  }
  if (typeof window.electronAPI !== 'undefined' && window.electronAPI.imap) {
    const config = await window.electronAPI.imap.getConfig();
    const accountKey = config && config.host && config.user ? config.host + '::' + config.user : '';
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
