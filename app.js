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

// --- IMAP cache (IndexedDB) ---
const IMAP_CACHE_DB_NAME = 'butter-mail-imap-cache';
const IMAP_CACHE_STORE = 'cache';

function openImapCacheDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IMAP_CACHE_DB_NAME, 1);
    req.onerror = () => reject(req.error);
    req.onsuccess = () => resolve(req.result);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
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
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IMAP_CACHE_STORE, 'readonly');
      const store = tx.objectStore(IMAP_CACHE_STORE);
      const req = store.get(accountKey);
      req.onsuccess = () => {
        const row = req.result;
        resolve(row && Array.isArray(row.emails) ? row.emails : []);
      };
      req.onerror = () => reject(req.error);
    });
  } catch {
    return [];
  }
}

async function setCachedEmails(accountKey, emails) {
  if (!accountKey || !Array.isArray(emails)) return;
  try {
    const db = await openImapCacheDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IMAP_CACHE_STORE, 'readwrite');
      const store = tx.objectStore(IMAP_CACHE_STORE);
      store.put({ accountKey, emails, lastSynced: Date.now() });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (e) {
    console.warn('[butter-mail] imap cache write failed:', e);
  }
}

async function getImapCacheLastSynced(accountKey) {
  if (!accountKey) return null;
  try {
    const db = await openImapCacheDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IMAP_CACHE_STORE, 'readonly');
      const req = tx.objectStore(IMAP_CACHE_STORE).get(accountKey);
      req.onsuccess = () => resolve(req.result ? req.result.lastSynced : null);
      req.onerror = () => reject(req.error);
    });
  } catch {
    return null;
  }
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
let selectedEmail = null;
let timelineViewInList = false;
const expandedThreads = new Set();
let isFetchingFromImap = false;
let isFetchingMore = false;
let imapInboxHasMore = true;

function updateLoadMoreButton() {
  const btn = document.getElementById('load-more-btn');
  if (!btn) return;
  const hasInbox = imapEmails.some((e) => e.mailbox === 'INBOX');
  const show = typeof window.electronAPI !== 'undefined' && hasInbox && imapInboxHasMore && !isFetchingMore;
  btn.style.display = show ? '' : 'none';
  btn.disabled = isFetchingMore;
  btn.textContent = isFetchingMore ? 'Loading…' : 'Load more';
}

function getFilteredEmails() {
  let emails = searchResults !== null ? searchResults : getAllEmails();
  if (currentFilter === 'all') return emails;
  if (currentFilter.startsWith('prompt-')) {
    const slug = currentFilter.slice(7);
    const clusters = getPromptClusters();
    const cluster = clusters[slug];
    if (!cluster) return [];
    const threshold = cluster.threshold != null ? cluster.threshold : 0.3;
    let idSet;
    if (cluster.scored && Array.isArray(cluster.scored)) {
      const overrides = cluster.overrides || {};
      const inCluster = (s) => overrides[s.id] === true || (overrides[s.id] !== false && s.sim >= threshold);
      idSet = new Set(cluster.scored.filter(inCluster).map((s) => s.id));
    } else if (cluster.emailIds) {
      idSet = new Set(cluster.emailIds);
    } else {
      return [];
    }
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
  return cats.meta[catId] && cats.meta[catId].color ? cats.meta[catId].color : '#B8952E';
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

  if (typeof ensureTimelineThreadHeaders === 'function') {
    await ensureTimelineThreadHeaders();
  }

  const progressEl = document.getElementById('progress-text');
  const btn = document.getElementById('compute-embeddings-btn');
  if (progressEl) {
    progressEl.textContent = 'Loading model...';
  }
  if (btn) {
    btn.disabled = true;
  }
  console.log('[butter-mail] computeEmbeddings: loading model...');

  if (window.electronAPI.embeddings.onProgress) {
    window.electronAPI.embeddings.onProgress((p) => {
      const message = `${p.current} / ${p.total} ${p.message || ''}`;
      if (progressEl) {
        progressEl.textContent = message;
      }
      console.log('[butter-mail] computeEmbeddings progress:', message);
    });
  }

  try {
    console.log('[butter-mail] computeEmbeddings: computing embeddings for', emails.length, 'emails');
    const res = await window.electronAPI.embeddings.compute(emails);
    if (!res.ok) throw new Error(res.error || 'Failed');
    saveEmbeddings(res.embeddings);
    console.log('[butter-mail] computeEmbeddings: embeddings computed and saved for', Object.keys(res.embeddings || {}).length, 'emails');

    const allEmailIds = Object.keys(res.embeddings);
    if (allEmailIds.length > 0) {
      if (progressEl) {
        progressEl.textContent = 'Running PCA...';
      }
      console.log('[butter-mail] PCA (graph view): running for', allEmailIds.length, 'emails');
      const pcaRes = await window.electronAPI.embeddings.pca(res.embeddings, allEmailIds);
      if (pcaRes.ok && pcaRes.points) {
        pcaPoints = pcaRes.points;
        savePcaModel(pcaRes.model);
        savePcaPoints(pcaPoints);
        console.log('[butter-mail] PCA (graph view): complete.');
      } else {
        console.warn('[butter-mail] PCA (graph view): did not return points.');
      }
      const { repIds, repToMembers } = buildThreadRepsForClustering(emails, res.embeddings);
      if (repIds.length > 0) {
        if (progressEl) {
          progressEl.textContent = 'Clustering emails...';
        }
        console.log('[butter-mail] clustering: starting for', repIds.length, 'thread representatives');
        const clusterRes = await window.electronAPI.embeddings.cluster(res.embeddings, repIds);
        if (clusterRes.ok) {
          const expanded = expandClusterAssignmentsToThreads(clusterRes, repToMembers, res.embeddings);
          saveCategories(expanded);
          console.log('[butter-mail] clustering: complete.');
        } else {
          console.warn('[butter-mail] clustering: did not return ok.');
        }
      }
    }
    if (progressEl) {
      progressEl.textContent = 'Done.';
    }
    console.log('[butter-mail] computeEmbeddings: done.');
    updateSubtabBar();
    refreshCurrentView();
  } catch (err) {
    if (progressEl) {
      progressEl.textContent = '';
    }
    console.error('[butter-mail] computeEmbeddings error:', err);
    alert('Error: ' + (err.message || String(err)));
  } finally {
    if (btn) {
      btn.disabled = false;
    }
    setTimeout(() => {
      if (progressEl) {
        progressEl.textContent = '';
      }
    }, 2000);
    console.log('[butter-mail] computeEmbeddings: finished cleanup.');
  }
}

// --- Cluster tabs in sidebar ---
let draggingCluster = null;
let pendingClusterRename = null;

function updateSubtabBar() {
  const listEl = document.getElementById('subtab-bar');
  if (!listEl) return;
  const cats = getCategories();
  const promptClusters = getPromptClusters();
  const promptEntries = Object.entries(promptClusters || {}).sort((a, b) => {
    const [slugA, clusterA] = a;
    const [slugB, clusterB] = b;
    const orderA = typeof clusterA.order === 'number' ? clusterA.order : Number.MAX_SAFE_INTEGER;
    const orderB = typeof clusterB.order === 'number' ? clusterB.order : Number.MAX_SAFE_INTEGER;
    if (orderA !== orderB) return orderA - orderB;
    const ca = clusterA.createdAt || '';
    const cb = clusterB.createdAt || '';
    if (ca && cb && ca !== cb) return ca.localeCompare(cb);
    return slugA.localeCompare(slugB);
  });
  const catIds = Object.keys(cats.meta || {}).filter((k) => k !== 'noise').sort((a, b) => {
    const ma = cats.meta[a] || {};
    const mb = cats.meta[b] || {};
    const orderA = typeof ma.order === 'number' ? ma.order : Number.MAX_SAFE_INTEGER;
    const orderB = typeof mb.order === 'number' ? mb.order : Number.MAX_SAFE_INTEGER;
    if (orderA !== orderB) return orderA - orderB;
    return a.localeCompare(b);
  });

  const activeTab = currentFilter;
  let html = '';
  html += '<div class="categories-option categories-option-all' + (activeTab === 'all' ? ' active' : '') + '" data-subtab="all" role="option"><span class="categories-option-label">all</span></div>';

  // User-defined prompt clusters (initially rendered first; drag-and-drop can mix)
  promptEntries.forEach(([slug, cluster]) => {
    const label = cluster && cluster.label ? cluster.label : slug;
    const tabId = 'prompt-' + slug;
    const isActive = activeTab === tabId;
    html += '<div class="categories-option categories-option-prompt' + (isActive ? ' active' : '') + '" data-subtab="' + escapeHtml(tabId) + '" data-prompt-slug="' + escapeHtml(slug) + '" role="option">' +
      '<span class="categories-option-label">' + escapeHtml(label) + '</span>' +
      '<div class="categories-option-actions">' +
      '<button type="button" class="categories-option-handle" title="Drag to reorder" aria-label="Drag to reorder">⋮⋮</button>' +
      '<span class="categories-option-type categories-option-type-user">user</span>' +
      '<button type="button" class="categories-option-rename" data-rename-prompt="' + escapeHtml(slug) + '" title="Rename cluster" aria-label="Rename cluster">&#9998;</button>' +
      '<button type="button" class="categories-option-delete" data-delete-prompt="' + escapeHtml(slug) + '" title="Delete cluster" aria-label="Delete cluster">&#215;</button>' +
      '</div>' +
      '</div>';
  });

  // Auto DBSCAN clusters after user-defined
  catIds.forEach((cid) => {
    const meta = cats.meta[cid];
    const name = meta && meta.name ? meta.name : cid;
    const color = meta && meta.color ? meta.color : '#B8952E';
    const isActive = activeTab === cid;
    html += '<div class="categories-option subtab-btn-category' + (isActive ? ' active' : '') + '" data-subtab="' + escapeHtml(cid) + '" data-category-id="' + escapeHtml(cid) + '" role="option" style="--cat-color:' + escapeHtml(color) + '">' +
      '<span class="categories-option-swatch" style="background-color:' + escapeHtml(color) + '"></span>' +
      '<span class="categories-option-label">' + escapeHtml(name) + '</span>' +
      '<div class="categories-option-actions">' +
      '<button type="button" class="categories-option-handle" title="Drag to reorder" aria-label="Drag to reorder">⋮⋮</button>' +
      '<span class="categories-option-type categories-option-type-auto">auto</span>' +
      '<button type="button" class="categories-option-rename" data-rename-category="' + escapeHtml(cid) + '" title="Rename category" aria-label="Rename category">&#9998;</button>' +
      '<button type="button" class="categories-option-delete" data-delete-category="' + escapeHtml(cid) + '" title="Delete category" aria-label="Delete category">&#215;</button>' +
      '</div>' +
      '</div>';
  });
  listEl.innerHTML = html;

  catIds.forEach((cid) => {
    const row = listEl.querySelector('[data-subtab="' + cid + '"]');
    if (row) row.addEventListener('dblclick', () => renameCategory(cid));
  });

  listEl.querySelectorAll('.categories-option-delete[data-delete-category]').forEach((delBtn) => {
    delBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      deleteCategory(delBtn.dataset.deleteCategory);
    });
  });
  listEl.querySelectorAll('.categories-option-delete[data-delete-prompt]').forEach((delBtn) => {
    delBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      deletePromptCluster(delBtn.dataset.deletePrompt);
    });
  });

  listEl.querySelectorAll('.categories-option-move[data-prompt-slug]').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const slug = btn.dataset.promptSlug;
      const dir = btn.dataset.move === 'up' ? 'up' : 'down';
      movePromptCluster(slug, dir);
    });
  });

  listEl.querySelectorAll('.categories-option-rename[data-rename-prompt]').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const slug = btn.dataset.renamePrompt;
      renamePromptCluster(slug);
    });
  });

  listEl.querySelectorAll('.categories-option-rename[data-rename-category]').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const catId = btn.dataset.renameCategory;
      renameCategory(catId);
    });
  });

  listEl.querySelectorAll('.categories-option').forEach((row) => {
    row.addEventListener('click', (e) => {
      if (e.target.closest('.categories-option-delete')) return;
      listEl.querySelectorAll('.categories-option').forEach((x) => x.classList.remove('active'));
      row.classList.add('active');
      currentFilter = row.dataset.subtab;
      refreshCurrentView();
    });
    if (row.dataset.subtab && row.dataset.subtab.startsWith('prompt-')) {
      const slug = row.dataset.subtab.slice(7);
      row.addEventListener('dblclick', (e) => {
        if (e.target.closest('.categories-option-delete')) return;
        e.preventDefault();
        openClusterThresholdModalForEdit(slug);
      });
    }
  });

  setupClusterDragAndDrop(listEl);
}

function setupClusterDragAndDrop(listEl) {
  listEl.querySelectorAll('.categories-option').forEach((row) => {
    const subtab = row.dataset.subtab;
    if (!subtab || subtab === 'all') return;
    row.setAttribute('draggable', 'true');
    row.addEventListener('dragstart', (e) => {
      draggingCluster = { subtab };
      row.classList.add('dragging');
      if (e.dataTransfer) {
        e.dataTransfer.effectAllowed = 'move';
      }
    });
    row.addEventListener('dragend', () => {
      draggingCluster = null;
      row.classList.remove('dragging');
      listEl.querySelectorAll('.categories-option').forEach((r) => r.classList.remove('drag-over'));
    });
    row.addEventListener('dragover', (e) => {
      if (!draggingCluster) return;
      e.preventDefault();
      if (e.dataTransfer) {
        e.dataTransfer.dropEffect = 'move';
      }
      listEl.querySelectorAll('.categories-option').forEach((r) => r.classList.remove('drag-over'));
      row.classList.add('drag-over');
    });
    row.addEventListener('drop', (e) => {
      if (!draggingCluster) return;
      e.preventDefault();
      row.classList.remove('drag-over');
      const fromSubtab = draggingCluster.subtab;
      const toSubtab = subtab;
      if (fromSubtab && toSubtab && fromSubtab !== toSubtab) {
        reorderClusters(fromSubtab, toSubtab);
      }
      draggingCluster = null;
    });
  });
}

function setupClusterRenameModal() {
  const overlay = document.getElementById('cluster-rename-overlay');
  const form = document.getElementById('cluster-rename-form');
  const input = document.getElementById('cluster-rename-input');
  const cancelBtn = document.getElementById('cluster-rename-cancel');
  const closeBtn = document.getElementById('cluster-rename-close');
  const titleEl = document.getElementById('cluster-rename-title');
  if (!overlay || !form || !input || !cancelBtn || !closeBtn || !titleEl) return;

  function close() {
    overlay.classList.add('hidden');
    pendingClusterRename = null;
    form.reset();
  }

  cancelBtn.onclick = close;
  closeBtn.onclick = close;
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) close();
  });

  form.addEventListener('submit', (e) => {
    e.preventDefault();
    if (!pendingClusterRename) {
      close();
      return;
    }
    const name = input.value.trim();
    if (!name) {
      close();
      return;
    }
    if (pendingClusterRename.type === 'prompt') {
      const clusters = getPromptClusters();
      const cluster = clusters[pendingClusterRename.id];
      if (cluster) {
        cluster.label = name;
        savePromptClusters(clusters);
      }
    } else if (pendingClusterRename.type === 'category') {
      const cats = getCategories();
      if (cats.meta[pendingClusterRename.id]) {
        cats.meta[pendingClusterRename.id].name = name;
        saveCategories(cats);
      }
    }
    close();
    updateSubtabBar();
    refreshCurrentView();
  });
}

function openClusterRenameModal(type, id, currentName) {
  const overlay = document.getElementById('cluster-rename-overlay');
  const input = document.getElementById('cluster-rename-input');
  const titleEl = document.getElementById('cluster-rename-title');
  if (!overlay || !input || !titleEl) return;
  pendingClusterRename = { type, id };
  titleEl.textContent = type === 'prompt' ? 'Rename cluster' : 'Rename category';
  input.value = currentName || '';
  overlay.classList.remove('hidden');
  setTimeout(() => {
    input.focus();
    input.select();
  }, 0);
}

function renameCategory(catId) {
  const cats = getCategories();
  if (!cats.meta[catId]) return;
  const currentName = cats.meta[catId].name || catId;
  openClusterRenameModal('category', catId, currentName);
}

function deleteCategory(catId) {
  const cats = getCategories();
  if (!cats.meta[catId]) return;
  const name = cats.meta[catId].name || catId;
  if (!confirm('Delete category "' + name + '"? Emails will move to Uncategorized.')) return;
  delete cats.meta[catId];
  Object.keys(cats.assignments).forEach((emailId) => {
    if (cats.assignments[emailId] === catId) cats.assignments[emailId] = 'noise';
  });
  if (!cats.meta['noise']) cats.meta['noise'] = { name: 'Uncategorized', color: '#999' };
  saveCategories(cats);
  if (currentFilter === catId) currentFilter = 'all';
  updateSubtabBar();
  refreshCurrentView();
}

function reorderCategoryClusters(fromId, toId) {
  const cats = getCategories();
  const meta = cats.meta || {};
  if (!meta[fromId] || !meta[toId]) return;
  const ids = Object.keys(meta).filter((k) => k !== 'noise').sort((a, b) => {
    const ma = meta[a] || {};
    const mb = meta[b] || {};
    const orderA = typeof ma.order === 'number' ? ma.order : Number.MAX_SAFE_INTEGER;
    const orderB = typeof mb.order === 'number' ? mb.order : Number.MAX_SAFE_INTEGER;
    if (orderA !== orderB) return orderA - orderB;
    return a.localeCompare(b);
  });
  const fromIndex = ids.indexOf(fromId);
  const toIndex = ids.indexOf(toId);
  if (fromIndex === -1 || toIndex === -1) return;
  const [item] = ids.splice(fromIndex, 1);
  const insertIndex = fromIndex < toIndex ? toIndex : toIndex;
  ids.splice(insertIndex, 0, item);
  ids.forEach((id, idx) => {
    if (meta[id]) meta[id].order = idx;
  });
  saveCategories(cats);
  updateSubtabBar();
}

function reorderClusters(fromSubtab, toSubtab) {
  const cats = getCategories();
  const promptClusters = getPromptClusters();
  const entries = [];

  Object.entries(promptClusters || {}).forEach(([slug, cluster]) => {
    entries.push({
      type: 'prompt',
      subtab: 'prompt-' + slug,
      id: slug,
      order: typeof cluster.order === 'number' ? cluster.order : Number.MAX_SAFE_INTEGER,
      createdAt: cluster.createdAt || ''
    });
  });

  Object.keys(cats.meta || {}).forEach((cid) => {
    if (cid === 'noise') return;
    const meta = cats.meta[cid] || {};
    entries.push({
      type: 'category',
      subtab: cid,
      id: cid,
      order: typeof meta.order === 'number' ? meta.order : Number.MAX_SAFE_INTEGER,
      createdAt: meta.createdAt || ''
    });
  });

  if (entries.length === 0) return;

  entries.sort((a, b) => {
    if (a.order !== b.order) return a.order - b.order;
    if (a.createdAt && b.createdAt && a.createdAt !== b.createdAt) return a.createdAt.localeCompare(b.createdAt);
    return String(a.subtab).localeCompare(String(b.subtab));
  });

  const fromIndex = entries.findIndex((e) => e.subtab === fromSubtab);
  const toIndex = entries.findIndex((e) => e.subtab === toSubtab);
  if (fromIndex === -1 || toIndex === -1) return;

  const [item] = entries.splice(fromIndex, 1);
  const insertIndex = fromIndex < toIndex ? toIndex : toIndex;
  entries.splice(insertIndex, 0, item);

  entries.forEach((entry, idx) => {
    if (entry.type === 'prompt') {
      if (promptClusters[entry.id]) {
        promptClusters[entry.id].order = idx;
      }
    } else if (entry.type === 'category') {
      if (!cats.meta[entry.id]) cats.meta[entry.id] = {};
      cats.meta[entry.id].order = idx;
    }
  });

  savePromptClusters(promptClusters);
  saveCategories(cats);
  updateSubtabBar();
}

function deletePromptCluster(slug) {
  const clusters = getPromptClusters();
  if (!clusters[slug]) return;
  const label = clusters[slug].label || slug;
  if (!confirm('Delete cluster "' + label + '"?')) return;
  delete clusters[slug];
  savePromptClusters(clusters);
  if (currentFilter === 'prompt-' + slug) currentFilter = 'all';
  updateSubtabBar();
  refreshCurrentView();
}

function renamePromptCluster(slug) {
  const clusters = getPromptClusters();
  const cluster = clusters[slug];
  if (!cluster) return;
  const currentLabel = cluster.label || slug;
  openClusterRenameModal('prompt', slug, currentLabel);
}

function movePromptCluster(slug, direction) {
  const clusters = getPromptClusters();
  const entries = Object.entries(clusters || {});
  if (!entries.length || !clusters[slug]) return;
  entries.sort((a, b) => {
    const [slugA, clusterA] = a;
    const [slugB, clusterB] = b;
    const orderA = typeof clusterA.order === 'number' ? clusterA.order : Number.MAX_SAFE_INTEGER;
    const orderB = typeof clusterB.order === 'number' ? clusterB.order : Number.MAX_SAFE_INTEGER;
    if (orderA !== orderB) return orderA - orderB;
    const ca = clusterA.createdAt || '';
    const cb = clusterB.createdAt || '';
    if (ca && cb && ca !== cb) return ca.localeCompare(cb);
    return slugA.localeCompare(slugB);
  });
  const index = entries.findIndex(([s]) => s === slug);
  if (index === -1) return;
  const targetIndex = direction === 'up' ? index - 1 : index + 1;
  if (targetIndex < 0 || targetIndex >= entries.length) return;
  const [item] = entries.splice(index, 1);
  entries.splice(targetIndex, 0, item);
  entries.forEach(([s, cluster], idx) => {
    cluster.order = idx;
  });
  savePromptClusters(clusters);
  updateSubtabBar();
}

// --- View switching ---
let timelineHeadersLoaded = false;

async function ensureTimelineThreadHeaders() {
  if (typeof window.electronAPI === 'undefined' || timelineHeadersLoaded) return;
  const imapWithoutThread = imapEmails.filter(
    (e) => !e.messageId && e.uid && (!e.mailbox || e.mailbox === 'INBOX')
  );
  if (imapWithoutThread.length === 0) {
    timelineHeadersLoaded = true;
    return;
  }
  const uids = imapWithoutThread.map((e) => e.uid).filter(Boolean);
  console.log('[butter-mail] timeline: fetching thread headers for', uids.length, 'emails');
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
    console.log('[butter-mail] timeline: thread headers loaded for', Object.keys(res.headers).length, 'emails');
  } catch (_) {}
}

function buildThreadRepsForClustering(allEmails, embeddings) {
  const emails = allEmails.filter((e) => embeddings[e.id]);
  if (emails.length === 0) {
    return { repIds: [], repToMembers: {} };
  }

  let threads = null;
  if (window.TimelineView && typeof window.TimelineView.buildThreads === 'function') {
    try {
      threads = window.TimelineView.buildThreads(emails);
    } catch (err) {
      console.warn('[butter-mail] timeline: buildThreads failed for clustering:', err);
      threads = null;
    }
  }

  if (!threads || !Array.isArray(threads) || threads.length === 0) {
    const repIdsFallback = emails.map((e) => e.id);
    const repToMembersFallback = {};
    repIdsFallback.forEach((id) => {
      repToMembersFallback[id] = [id];
    });
    console.log('[butter-mail] clustering: using per-email reps (no threads available) for', emails.length, 'emails');
    return { repIds: repIdsFallback, repToMembers: repToMembersFallback };
  }

  const repIds = [];
  const repToMembers = {};
  const memberSet = new Set();

  threads.forEach((thread) => {
    if (!thread || thread.length === 0) return;
    const members = thread.filter((e) => embeddings[e.id]);
    if (members.length === 0) return;
    const rep = members[members.length - 1]; // latest email in thread (buildThreads sorts by date asc)
    const repId = rep.id;
    const memberIds = members.map((e) => e.id);
    repIds.push(repId);
    repToMembers[repId] = memberIds;
    memberIds.forEach((id) => memberSet.add(id));
  });

  emails.forEach((e) => {
    if (!memberSet.has(e.id)) {
      repIds.push(e.id);
      repToMembers[e.id] = [e.id];
      memberSet.add(e.id);
    }
  });

  console.log('[butter-mail] clustering: using', repIds.length, 'thread representatives for', emails.length, 'emails');
  return { repIds, repToMembers };
}

function expandClusterAssignmentsToThreads(clusterRes, repToMembers, embeddings) {
  const assignments = {};
  const meta = { ...(clusterRes.meta || {}) };
  const srcAssignments = clusterRes.assignments || {};

  Object.keys(repToMembers).forEach((repId) => {
    const catId = srcAssignments[repId] || 'noise';
    const members = repToMembers[repId] || [];
    members.forEach((emailId) => {
      assignments[emailId] = catId;
    });
  });

  Object.keys(embeddings || {}).forEach((emailId) => {
    if (!assignments[emailId]) {
      assignments[emailId] = 'noise';
    }
  });

  if (!meta.noise) {
    meta.noise = { name: 'Uncategorized', color: '#999' };
  }

  return { assignments, meta };
}

function refreshCurrentView() {
  const emails = getFilteredEmails();
  if (currentView === 'list') {
    renderEmailList(emails);
  } else if (currentView === 'graph') {
    renderGraphView(emails);
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
  const pointCount = Object.keys(points).length;
  console.log('[butter-mail] graph view: rendering', pointCount, 'points (PCA).');
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

// --- Fetch IMAP ---
async function fetchFromImap() {
  if (typeof window.electronAPI === 'undefined') return;
  if (isFetchingFromImap) {
    console.log('[butter-mail] refresh: skipped (already fetching)');
    return;
  }
  isFetchingFromImap = true;
  const btn = document.getElementById('refresh-btn');
  console.log('[butter-mail] refresh: start');
  if (btn) {
    btn.classList.add('loading');
    btn.disabled = true;
  } else {
    console.warn('[butter-mail] refresh: refresh button element not found');
  }
  try {
    const config = await window.electronAPI.imap.getConfig();
    const accountKey = config && config.host && config.user ? config.host + '::' + config.user : null;
    console.log('[butter-mail] refresh: calling main imap.fetch(limit=150)');
    const result = await window.electronAPI.imap.fetch(150);
    if (result.ok) {
      console.log('[butter-mail] refresh: ok. emails:', Array.isArray(result.emails) ? result.emails.length : 0);
      imapEmails = result.emails;
      timelineHeadersLoaded = false;
      if (accountKey) await setCachedEmails(accountKey, result.emails);
    } else if (!result.error.includes('not configured')) {
      const message = result.error || 'Unknown';
      console.error('[butter-mail] IMAP refresh failed:', message);
      alert('IMAP refresh failed: ' + message);
    }
    refreshCurrentView();
    updateSubtabBar();
    updateLoadMoreButton();
  } finally {
    isFetchingFromImap = false;
    console.log('[butter-mail] refresh: end');
    if (btn) {
      btn.classList.remove('loading');
      btn.disabled = false;
    }
    updateLoadMoreButton();
  }
}

// --- Re-cluster (DBSCAN only; keeps prompt clusters) ---
async function recluster() {
  if (typeof window.electronAPI === 'undefined') {
    alert('Re-cluster requires the Electron app. Run: npm start');
    return;
  }
  const embeddings = getEmbeddings();
  const emailIds = Object.keys(embeddings);
  if (emailIds.length === 0) {
    alert('No embeddings. Run "compute embeddings" first.');
    return;
  }
  const allEmails = getAllEmails();
  if (typeof ensureTimelineThreadHeaders === 'function') {
    await ensureTimelineThreadHeaders();
  }
  const { repIds, repToMembers } = buildThreadRepsForClustering(allEmails, embeddings);
  const btn = document.getElementById('recluster-btn');
  const progressEl = document.getElementById('progress-text');
  if (btn) btn.disabled = true;
  if (progressEl) progressEl.textContent = 'Re-clustering...';
  console.log('[butter-mail] recluster: starting for', repIds.length, 'thread representatives');
  try {
    const clusterRes = await window.electronAPI.embeddings.cluster(embeddings, repIds);
    if (!clusterRes.ok) throw new Error(clusterRes.error || 'Re-cluster failed');
    // Full replace: previous DBSCAN categories are removed; prompt clusters (separate store) are untouched
    const expanded = expandClusterAssignmentsToThreads(clusterRes, repToMembers, embeddings);
    saveCategories(expanded);
    currentFilter = 'all';
    console.log('[butter-mail] recluster: complete.');
    updateSubtabBar();
    refreshCurrentView();
    if (progressEl) progressEl.textContent = 'Done.';
  } catch (err) {
    if (progressEl) progressEl.textContent = '';
    console.error('[butter-mail] recluster error:', err);
    alert('Error: ' + (err.message || String(err)));
  } finally {
    if (btn) btn.disabled = false;
    setTimeout(() => { if (progressEl) progressEl.textContent = ''; }, 2000);
  }
}

document.getElementById('refresh-btn').addEventListener('click', () => {
  console.log('[butter-mail] refresh button: clicked');
  fetchFromImap();
});
const loadMoreBtn = document.getElementById('load-more-btn');
if (loadMoreBtn) {
  loadMoreBtn.addEventListener('click', () => loadMoreImapEmails());
}
document.getElementById('compute-embeddings-btn').addEventListener('click', computeEmbeddings);
document.getElementById('recluster-btn').addEventListener('click', recluster);

let pendingClusterScored = null;
let pendingClusterPrompt = '';

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
  progressEl.textContent = 'Loading similarities...';
  try {
    console.log('[butter-mail] promptCluster: starting scored prompt cluster for', emailIds.length, 'emails with prompt:', prompt);
    const res = await window.electronAPI.embeddings.promptClusterScored(prompt, embeddings, emailIds);
    if (!res.ok) throw new Error(res.error || 'Failed');
    const scored = res.scored || [];
    console.log('[butter-mail] promptCluster: received', scored.length, 'scored results for prompt:', prompt);
    progressEl.textContent = '';
    btn.disabled = false;
    pendingClusterScored = scored;
    pendingClusterPrompt = prompt;
    openClusterThresholdModal(prompt, scored);
  } catch (err) {
    progressEl.textContent = '';
    console.error('[butter-mail] promptCluster error:', err);
    alert('Error: ' + (err.message || String(err)));
    btn.disabled = false;
  }
}

function clusterIncluded(s, threshold, overrides) {
  return overrides[s.id] === true || (overrides[s.id] !== false && s.sim >= threshold);
}

function renderClusterResultsList(resultsEl, scored, threshold, overrides, emailsById, countEl, onRowClick) {
  if (!resultsEl) return;
  if (scored.length === 0) {
    resultsEl.innerHTML = '<p class="settings-hint" style="margin: 0.5rem 0.75rem;">No matches.</p>';
    if (countEl) countEl.textContent = '0 email(s) in cluster';
    return;
  }
  let inCount = 0;
  resultsEl.innerHTML = scored.map((s) => {
    const inCluster = clusterIncluded(s, threshold, overrides);
    if (inCluster) inCount++;
    const email = emailsById[s.id];
    const subject = email && (email.subject || '').trim() ? email.subject : '(no subject)';
    const rowClass = inCluster ? 'cluster-result-row in-cluster' : 'cluster-result-row not-in-cluster';
    return '<div class="' + rowClass + '" data-id="' + escapeHtml(s.id) + '" title="Click to include/exclude">' +
      '<span class="cluster-result-subject" title="' + escapeHtml(subject) + '">' + escapeHtml(subject) + '</span>' +
      '<span class="cluster-result-sim">' + s.sim.toFixed(2) + '</span>' +
      '</div>';
  }).join('');
  if (countEl) countEl.textContent = inCount + ' email(s) in cluster';
  resultsEl.querySelectorAll('.cluster-result-row').forEach((row) => {
    row.addEventListener('click', () => { onRowClick(row.dataset.id); });
  });
}

function openClusterThresholdModal(prompt, scored) {
  const overlay = document.getElementById('cluster-threshold-overlay');
  const titleEl = document.getElementById('cluster-threshold-title');
  const resultsEl = document.getElementById('cluster-threshold-results');
  const sliderEl = document.getElementById('cluster-threshold-slider');
  const valueEl = document.getElementById('cluster-threshold-value');
  const countEl = document.getElementById('cluster-threshold-count');
  if (!overlay || !sliderEl) return;
  titleEl.textContent = 'Create cluster: "' + prompt + '"';
  const emailsById = {};
  getAllEmails().forEach((e) => { emailsById[e.id] = e; });
  const overrides = {};
  function getThreshold() { return parseFloat(sliderEl.value, 10); }
  function refresh() {
    const t = getThreshold();
    valueEl.textContent = t.toFixed(2);
    renderClusterResultsList(resultsEl, scored, t, overrides, emailsById, countEl, (id) => {
      const s = scored.find((x) => x.id === id);
      if (!s) return;
      const inCluster = clusterIncluded(s, t, overrides);
      overrides[id] = inCluster ? false : true;
      refresh();
    });
  }
  sliderEl.value = '0.3';
  sliderEl.addEventListener('input', refresh);
  refresh();
  overlay.classList.remove('hidden');

  function closeModal() {
    overlay.classList.add('hidden');
    sliderEl.removeEventListener('input', refresh);
  }

  document.getElementById('cluster-threshold-close').onclick = closeModal;
  document.getElementById('cluster-threshold-cancel').onclick = closeModal;
  overlay.onclick = (e) => { if (e.target === overlay) closeModal(); };

  document.getElementById('cluster-threshold-create').onclick = () => {
    const threshold = getThreshold();
    const slug = pendingClusterPrompt.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'cluster';
    const clusters = getPromptClusters();
    let finalSlug = slug;
    let idx = 0;
    while (clusters[finalSlug]) {
      idx++;
      finalSlug = slug + '-' + idx;
    }
    clusters[finalSlug] = {
      label: pendingClusterPrompt,
      threshold,
      scored: pendingClusterScored,
      overrides: Object.keys(overrides).length ? overrides : undefined,
      createdAt: new Date().toISOString()
    };
    savePromptClusters(clusters);
    const input = document.getElementById('prompt-cluster-input');
    if (input) input.value = '';
    pendingClusterScored = null;
    pendingClusterPrompt = '';
    closeModal();
    updateSubtabBar();
    refreshCurrentView();
  };
}

function openClusterThresholdModalForEdit(slug) {
  const clusters = getPromptClusters();
  const cluster = clusters[slug];
  if (!cluster || !cluster.scored || !Array.isArray(cluster.scored)) return;
  const overlay = document.getElementById('cluster-threshold-overlay');
  const titleEl = document.getElementById('cluster-threshold-title');
  const resultsEl = document.getElementById('cluster-threshold-results');
  const sliderEl = document.getElementById('cluster-threshold-slider');
  const valueEl = document.getElementById('cluster-threshold-value');
  const countEl = document.getElementById('cluster-threshold-count');
  const createBtn = document.getElementById('cluster-threshold-create');
  if (!overlay || !sliderEl) return;
  const label = cluster.label || slug;
  titleEl.textContent = 'Edit cluster: "' + label + '"';
  const emailsById = {};
  getAllEmails().forEach((e) => { emailsById[e.id] = e; });
  const overrides = cluster.overrides ? { ...cluster.overrides } : {};
  const scored = cluster.scored;
  function getThreshold() { return parseFloat(sliderEl.value, 10); }
  function refresh() {
    const t = getThreshold();
    valueEl.textContent = t.toFixed(2);
    renderClusterResultsList(resultsEl, scored, t, overrides, emailsById, countEl, (id) => {
      const s = scored.find((x) => x.id === id);
      if (!s) return;
      const inCluster = clusterIncluded(s, t, overrides);
      overrides[id] = inCluster ? false : true;
      refresh();
    });
  }
  const currentThreshold = cluster.threshold != null ? cluster.threshold : 0.3;
  sliderEl.value = String(currentThreshold);
  sliderEl.addEventListener('input', refresh);
  refresh();
  overlay.classList.remove('hidden');
  createBtn.textContent = 'Save';

  function closeModal() {
    overlay.classList.add('hidden');
    sliderEl.removeEventListener('input', refresh);
    createBtn.textContent = 'Create cluster';
  }

  document.getElementById('cluster-threshold-close').onclick = closeModal;
  document.getElementById('cluster-threshold-cancel').onclick = closeModal;
  overlay.onclick = (e) => { if (e.target === overlay) closeModal(); };

  createBtn.onclick = () => {
    const threshold = getThreshold();
    cluster.threshold = threshold;
    cluster.overrides = Object.keys(overrides).length ? overrides : undefined;
    savePromptClusters(clusters);
    closeModal();
    refreshCurrentView();
  };
}

document.getElementById('prompt-cluster-btn').addEventListener('click', createPromptCluster);
document.getElementById('prompt-cluster-input').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') createPromptCluster();
});

// --- Virtual list (flat list only) ---
const VIRTUAL_LIST_ROW_HEIGHT = 60;
const VIRTUAL_LIST_THRESHOLD = 20;
const VIRTUAL_LIST_OVERSCAN = 5;

function buildFlatRowHtml(e, isSearchMode, cats, addVirtualClass) {
  const catColor = e.categoryId && cats.meta[e.categoryId] ? cats.meta[e.categoryId].color : '#B8952E';
  const catTitle = e.categoryId && cats.meta[e.categoryId] ? (cats.meta[e.categoryId].name || '') : '';
  const rankHtml = isSearchMode
    ? '<span class="email-row-rank">#' + String(e.searchRank) + '</span>' +
      '<span class="email-row-scores">' +
      'dense: ' + (e.denseScore != null ? Number(e.denseScore).toFixed(2) : '—') + ' | ' +
      'sparse: ' + (e.sparseScore != null ? Number(e.sparseScore).toFixed(2) : '—') +
      '</span>'
    : '';
  const virtualClass = addVirtualClass ? ' email-row-virtual' : '';
  return '<div class="email-row' + virtualClass + (isSearchMode ? ' email-row-search' : '') + '" data-id="' + escapeHtml(e.id) + '">' +
    '<span class="email-row-category-square" style="background-color:' + escapeHtml(catColor) + '" title="' + escapeHtml(catTitle) + '" aria-hidden></span>' +
    (isSearchMode ? rankHtml : '') +
    '<span class="email-row-subject" style="text-decoration-color:' + escapeHtml(catColor) + '">' + escapeHtml(e.subject) + '</span>' +
    '<span class="email-row-date">' + escapeHtml(formatDate(e.date)) + '</span>' +
    '</div>';
}

function attachVirtualRowClickHandlers(contentEl, listEl) {
  if (!contentEl || !listEl) return;
  contentEl.querySelectorAll('.email-row').forEach((row) => {
    row.addEventListener('click', () => {
      const email = getAllEmails().find((x) => x.id === row.dataset.id);
      if (email) {
        selectedEmail = email;
        if (currentView === 'list') {
          updateInlineDetail(email);
          listEl.querySelectorAll('.email-row').forEach((r) => r.classList.remove('selected'));
          row.classList.add('selected');
        } else {
          openEmailDetail(email);
        }
      }
    });
  });
  if (selectedEmail) {
    const sel = contentEl.querySelector('.email-row[data-id="' + escapeHtml(selectedEmail.id) + '"]');
    if (sel) sel.classList.add('selected');
  }
}

function updateVirtualListVisibleRows(listEl) {
  const data = listEl._virtualListData;
  if (!data || !data.wrapper || !data.contentEl || !data.topSpacer || !data.bottomSpacer) return;
  const { emails, isSearchMode, cats, totalRows } = data;
  const scrollTop = listEl.scrollTop;
  const containerHeight = listEl.clientHeight;
  const startIndex = Math.max(0, Math.floor(scrollTop / VIRTUAL_LIST_ROW_HEIGHT) - VIRTUAL_LIST_OVERSCAN);
  const visibleCount = Math.ceil(containerHeight / VIRTUAL_LIST_ROW_HEIGHT) + 2 * VIRTUAL_LIST_OVERSCAN;
  const endIndex = Math.min(totalRows, startIndex + visibleCount);
  // Skip DOM update if visible range unchanged (reduces scroll lag)
  if (data.lastStartIndex === startIndex && data.lastEndIndex === endIndex) return;
  data.lastStartIndex = startIndex;
  data.lastEndIndex = endIndex;
  data.topSpacer.style.height = (startIndex * VIRTUAL_LIST_ROW_HEIGHT) + 'px';
  data.bottomSpacer.style.height = ((totalRows - endIndex) * VIRTUAL_LIST_ROW_HEIGHT) + 'px';
  const slice = emails.slice(startIndex, endIndex);
  data.contentEl.innerHTML = slice.map((e) => buildFlatRowHtml(e, isSearchMode, cats, true)).join('');
  attachVirtualRowClickHandlers(data.contentEl, listEl);
}

// --- Render email list ---
function renderEmailList(emails) {
  const listEl = document.getElementById('email-list');
  if (!listEl) return;

  const isSearchMode = emails.length > 0 && typeof emails[0].searchRank === 'number';
  const sorted = isSearchMode
    ? [...emails].sort((a, b) => (a.searchRank || 0) - (b.searchRank || 0))
    : [...emails].sort((a, b) => {
        const da = new Date(a.date || 0);
        const db = new Date(b.date || 0);
        return db - da;
      });

  if (sorted.length === 0) {
    const hint = typeof window.electronAPI !== 'undefined'
      ? 'Click "refresh" to fetch from IMAP.'
      : 'Fetch from IMAP to get started.';
    listEl.innerHTML = '<p class="email-list-empty">No emails yet. ' + hint + '</p>';
    listEl._virtualListData = null;
    return;
  }

  const cats = getCategories();

  // If timeline view is enabled for the list, render one compact row per thread
  // with optional expansion (even when in search mode).
  const canUseTimelineThreads = timelineViewInList && window.TimelineView && typeof window.TimelineView.buildThreads === 'function';

  if (canUseTimelineThreads) {
    const emailsWithCat = getEmailsWithCategories(sorted);
    console.log('[butter-mail] timeline: building threads for', emailsWithCat.length, 'emails');
    const threads = window.TimelineView.buildThreads(emailsWithCat);
    console.log('[butter-mail] timeline: built', threads.length, 'threads, rendering list');

    const parts = [];
    threads.forEach((thread) => {
      if (!thread || thread.length === 0) return;
      const lastEmail = thread[thread.length - 1] || thread[0];
      const threadId = lastEmail.id;
      const hasMultiple = thread.length > 1;

      const containsSelected = selectedEmail && thread.some((e) => e.id === selectedEmail.id);
      if (containsSelected) {
        expandedThreads.add(threadId);
      }
      const isExpanded = hasMultiple && expandedThreads.has(threadId);

      const catMeta = lastEmail.categoryId && cats.meta[lastEmail.categoryId] ? cats.meta[lastEmail.categoryId] : null;
      const catColor = catMeta && catMeta.color ? catMeta.color : '#B8952E';
      const catTitle = catMeta && catMeta.name ? catMeta.name : '';

      const threadSubject = lastEmail.subject || '(no subject)';
      const threadDate = formatDate(lastEmail.date);

      parts.push(
        '<div class="email-thread-block" data-thread-id="' + escapeHtml(threadId) + '">'
        + '<div class="email-row email-row-thread' + (hasMultiple ? ' email-row-thread-stacked' : '') + '" data-thread-id="' + escapeHtml(threadId) + '" data-thread-size="' + String(thread.length) + '" data-latest-email-id="' + escapeHtml(lastEmail.id) + '">'
        + '<span class="email-row-category-square" style="background-color:' + escapeHtml(catColor) + '" title="' + escapeHtml(catTitle) + '" aria-hidden></span>'
        + '<span class="email-row-subject" style="text-decoration-color:' + escapeHtml(catColor) + '">' + escapeHtml(threadSubject) + '</span>'
        + '<span class="email-row-date">' + escapeHtml(threadDate) + '</span>'
        + (hasMultiple ? '<span class="thread-count-badge">+ ' + String(thread.length) + '</span>' : '')
        + '</div>'
      );

      if (isExpanded && hasMultiple) {
        parts.push('<div class="thread-email-list">');
        thread.forEach((email) => {
          const eCatMeta = email.categoryId && cats.meta[email.categoryId] ? cats.meta[email.categoryId] : null;
          const eCatColor = eCatMeta && eCatMeta.color ? eCatMeta.color : '#B8952E';
          const eCatTitle = eCatMeta && eCatMeta.name ? eCatMeta.name : '';
          const eSubject = email.subject || '(no subject)';
          const eDate = formatDate(email.date);
          const isSelected = selectedEmail && selectedEmail.id === email.id;

          parts.push(
            '<div class="email-row email-row-in-thread' + (isSelected ? ' selected' : '') + '" data-id="' + escapeHtml(email.id) + '" data-thread-id="' + escapeHtml(threadId) + '">'
            + '<span class="email-row-category-square" style="background-color:' + escapeHtml(eCatColor) + '" title="' + escapeHtml(eCatTitle) + '" aria-hidden></span>'
            + '<span class="email-row-subject" style="text-decoration-color:' + escapeHtml(eCatColor) + '">' + escapeHtml(eSubject) + '</span>'
            + '<span class="email-row-date">' + escapeHtml(eDate) + '</span>'
            + '</div>'
          );
        });
        parts.push('</div>');
      }

      parts.push('</div>');
    });

    listEl.innerHTML = parts.join('');

    // Thread header click: expand/collapse for multi-email threads, or open single-email thread.
    listEl.querySelectorAll('.email-row-thread').forEach((row) => {
      row.addEventListener('click', () => {
        const threadId = row.dataset.threadId;
        const size = Number(row.dataset.threadSize || '1');

        if (!threadId) return;

        if (size <= 1) {
          const latestId = row.dataset.latestEmailId;
          const email = getAllEmails().find((x) => x.id === latestId);
          if (email) {
            selectedEmail = email;
            updateInlineDetail(email);
            renderEmailList(getFilteredEmails());
          }
          return;
        }

        if (expandedThreads.has(threadId)) {
          expandedThreads.delete(threadId);
        } else {
          expandedThreads.add(threadId);
        }
        renderEmailList(getFilteredEmails());
      });
    });

    // Email row click inside expanded thread: open that email.
    listEl.querySelectorAll('.email-row-in-thread').forEach((row) => {
      row.addEventListener('click', (e) => {
        e.stopPropagation();
        const email = getAllEmails().find((x) => x.id === row.dataset.id);
        if (email) {
          selectedEmail = email;
          updateInlineDetail(email);
          renderEmailList(getFilteredEmails());
        }
      });
    });

    return;
  }

  // Default: flat list (optionally in search mode). Use virtual list when many rows.
  const emailsWithCat = getEmailsWithCategories(sorted);
  const totalRows = emailsWithCat.length;
  const useVirtualList = totalRows > VIRTUAL_LIST_THRESHOLD;

  if (useVirtualList) {
    listEl.innerHTML = '';
    const wrapper = document.createElement('div');
    wrapper.className = 'email-list-virtual-wrapper';
    wrapper.style.height = (totalRows * VIRTUAL_LIST_ROW_HEIGHT) + 'px';
    wrapper.style.position = 'relative';
    const topSpacer = document.createElement('div');
    topSpacer.className = 'email-list-virtual-spacer';
    topSpacer.style.height = '0';
    const contentEl = document.createElement('div');
    contentEl.className = 'email-list-virtual-content';
    const bottomSpacer = document.createElement('div');
    bottomSpacer.className = 'email-list-virtual-spacer';
    wrapper.appendChild(topSpacer);
    wrapper.appendChild(contentEl);
    wrapper.appendChild(bottomSpacer);
    listEl.appendChild(wrapper);
    listEl._virtualListData = {
      emails: emailsWithCat,
      isSearchMode,
      cats,
      totalRows,
      wrapper,
      topSpacer,
      contentEl,
      bottomSpacer,
      lastStartIndex: -1,
      lastEndIndex: -1
    };
    let scrollRaf = null;
    listEl.addEventListener('scroll', () => {
      if (scrollRaf) return;
      scrollRaf = requestAnimationFrame(() => {
        updateVirtualListVisibleRows(listEl);
        scrollRaf = null;
      });
    }, { passive: true });
    updateVirtualListVisibleRows(listEl);
    return;
  }

  // Non-virtual flat list (small list).
  listEl._virtualListData = null;
  listEl.innerHTML = emailsWithCat
    .map((e) => buildFlatRowHtml(e, isSearchMode, cats, false))
    .join('');

  listEl.querySelectorAll('.email-row').forEach((row) => {
    row.addEventListener('click', () => {
      const email = getAllEmails().find((x) => x.id === row.dataset.id);
      if (email) {
        selectedEmail = email;
        if (currentView === 'list') {
          updateInlineDetail(email);
          listEl.querySelectorAll('.email-row').forEach((r) => r.classList.remove('selected'));
          row.classList.add('selected');
        } else {
          openEmailDetail(email);
        }
      }
    });
  });

  if (selectedEmail) {
    const selectedRow = listEl.querySelector('.email-row[data-id="' + escapeHtml(selectedEmail.id) + '"]');
    if (selectedRow) selectedRow.classList.add('selected');
  }
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

function setInlineDetailContent(email, body, bodyIsHtml) {
  document.getElementById('email-detail-inline-subject').textContent = email.subject;
  document.getElementById('email-detail-inline-from').textContent = email.from || email.fromEmail || '';
  document.getElementById('email-detail-inline-to').textContent = email.toDisplay || email.to || '';
  document.getElementById('email-detail-inline-date').textContent = email.date ? new Date(email.date).toLocaleString() : '';
  document.getElementById('email-detail-inline-body').innerHTML = renderEmailBody(body || '', bodyIsHtml);
  document.getElementById('email-detail-inline-placeholder').classList.add('hidden');
  document.getElementById('email-detail-inline-content').classList.remove('hidden');
}

function clearInlineDetail() {
  document.getElementById('email-detail-inline-placeholder').classList.remove('hidden');
  document.getElementById('email-detail-inline-content').classList.add('hidden');
}

async function updateInlineDetail(email) {
  let body = email.body;
  let bodyIsHtml = email.bodyIsHtml || false;
  setInlineDetailContent(email, body, bodyIsHtml);
  if (!body && email.id && email.id.startsWith('imap-') && email.uid && typeof window.electronAPI !== 'undefined') {
    document.getElementById('email-detail-inline-body').textContent = 'Loading...';
    const result = await window.electronAPI.imap.fetchOne(email.uid);
    if (result.ok && result.email) {
      body = result.email.body;
      bodyIsHtml = result.email.bodyIsHtml || false;
      if (result.email.toDisplay) email.toDisplay = result.email.toDisplay;
    }
    setInlineDetailContent(email, body, bodyIsHtml);
  }
}

async function openEmailDetail(email) {
  if (currentView === 'list') {
    selectedEmail = email;
    await updateInlineDetail(email);
    renderEmailList(getFilteredEmails());
    return;
  }
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
    if (currentView !== 'list') {
      selectedEmail = null;
      clearInlineDetail();
    }
    refreshCurrentView();
  });
});

// --- List timeline view checkbox ---
const timelineCheckbox = document.getElementById('timeline-view-checkbox');
if (timelineCheckbox) {
  timelineCheckbox.addEventListener('change', () => {
    timelineViewInList = timelineCheckbox.checked;
    if (!timelineViewInList) {
      expandedThreads.clear();
    }
    if (timelineViewInList && typeof ensureTimelineThreadHeaders === 'function') {
      ensureTimelineThreadHeaders().then(() => {
        if (currentView === 'list' && timelineViewInList) {
          renderEmailList(getFilteredEmails());
        }
      });
    } else {
      renderEmailList(getFilteredEmails());
    }
  });
}

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

// --- Compose / Send ---
const composeOverlay = document.getElementById('compose-overlay');
const composeForm = document.getElementById('compose-form');
const composeStatus = document.getElementById('compose-status');

function openCompose(opts = {}) {
  document.getElementById('compose-to').value = opts.to || '';
  document.getElementById('compose-subject').value = opts.subject || '';
  document.getElementById('compose-body').value = opts.body || '';
  composeStatus.textContent = '';
  composeStatus.className = 'compose-status';
  composeOverlay.classList.remove('hidden');
}

function closeCompose() {
  composeOverlay.classList.add('hidden');
}

document.getElementById('compose-btn').addEventListener('click', () => {
  if (typeof window.electronAPI === 'undefined') {
    alert('Sending is available in the Electron app. Run: npm start');
    return;
  }
  openCompose();
});

document.getElementById('compose-close').addEventListener('click', closeCompose);
document.getElementById('compose-cancel').addEventListener('click', closeCompose);

composeOverlay.addEventListener('click', (e) => {
  if (e.target === composeOverlay) closeCompose();
});

document.getElementById('reply-inline-btn').addEventListener('click', () => {
  if (typeof window.electronAPI === 'undefined') {
    alert('Sending is available in the Electron app. Run: npm start');
    return;
  }
  const email = selectedEmail;
  if (!email) return;
  const replyTo = email.fromEmail || (email.from && email.from.match(/<([^>]+)>/) ? email.from.match(/<([^>]+)>/)[1] : email.from) || email.from || '';
  const replySubject = (email.subject || '').trim().replace(/^Re:\s*/i, '') ? 'Re: ' + (email.subject || '').trim() : 'Re: (no subject)';
  openCompose({ to: replyTo, subject: replySubject, body: '' });
});

composeForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  if (typeof window.electronAPI === 'undefined') return;
  const to = document.getElementById('compose-to').value.trim();
  const subject = document.getElementById('compose-subject').value.trim();
  const body = document.getElementById('compose-body').value.trim();
  const butterMailUrl = 'https://github.com/mehek-niwas/butter-mail';
  const signatureText = '\n\n- sent with butter mail ' + butterMailUrl;
  const text = body ? body + signatureText : signatureText.trim();

  const urlRe = /(https?:\/\/[^\s<]+)/g;
  const bodyHtml = escapeHtml(body)
    .replace(urlRe, (m) => '<a href="' + escapeHtml(m) + '">' + escapeHtml(m) + '</a>')
    .replace(/\n/g, '<br>');
  const signatureHtml = '- sent with <a href="' + butterMailUrl + '">butter mail</a>';
  const html = (body ? bodyHtml + '<br><br>' : '') + signatureHtml;
  composeStatus.textContent = 'Sending...';
  composeStatus.className = 'compose-status';
  const result = await window.electronAPI.smtp.send({ to, subject, text, html });
  if (result.ok) {
    composeStatus.textContent = 'Sent.';
    composeStatus.className = 'compose-status success';
    composeForm.reset();
    // After a successful send, refresh from IMAP so the new sent email appears
    try {
      if (typeof window.electronAPI !== 'undefined' && window.electronAPI.imap && typeof fetchFromImap === 'function') {
        await fetchFromImap();
      }
    } catch (err) {
      console.error('[butter-mail] auto-refresh after send failed:', err);
    }
    setTimeout(closeCompose, 1200);
  } else {
    composeStatus.textContent = result.error || 'Send failed.';
    composeStatus.className = 'compose-status error';
  }
});

// --- Initial load ---
setupClusterRenameModal();
updateSubtabBar();
updateLoadMoreButton();
refreshCurrentView();
if (typeof window.electronAPI !== 'undefined') {
  (async () => {
    try {
      const config = await window.electronAPI.imap.getConfig();
      const accountKey = config && config.host && config.user ? config.host + '::' + config.user : null;
      if (accountKey) {
        const cached = await getCachedEmails(accountKey);
        if (cached && cached.length > 0) {
          imapEmails = cached;
          refreshCurrentView();
          updateSubtabBar();
          updateLoadMoreButton();
        }
      }
    } catch (e) {
      console.warn('[butter-mail] imap cache load failed:', e);
    }
    fetchFromImap();
  })();
}

async function loadMoreImapEmails() {
  if (typeof window.electronAPI === 'undefined' || isFetchingMore) return;
  const inboxUids = imapEmails.filter((e) => e.mailbox === 'INBOX').map((e) => e.uid).filter((u) => u != null);
  if (inboxUids.length === 0) return;
  const minUid = Math.min(...inboxUids);
  isFetchingMore = true;
  updateLoadMoreButton();
  try {
    const result = await window.electronAPI.imap.fetchMore(100, minUid);
    if (result.ok && Array.isArray(result.emails) && result.emails.length > 0) {
      const existingIds = new Set(imapEmails.map((e) => e.id));
      const newEmails = result.emails.filter((e) => !existingIds.has(e.id));
      imapEmails = [...imapEmails, ...newEmails];
      const config = await window.electronAPI.imap.getConfig();
      const accountKey = config && config.host && config.user ? config.host + '::' + config.user : null;
      if (accountKey) await setCachedEmails(accountKey, imapEmails);
      refreshCurrentView();
    }
    if (result.hasMore === false) {
      imapInboxHasMore = false;
    }
  } catch (e) {
    console.warn('[butter-mail] load more failed:', e);
  } finally {
    isFetchingMore = false;
    updateLoadMoreButton();
  }
}
