/**
 * Hybrid search: BM25 (sparse) + embeddings (dense) + RRF.
 * Runs in Electron main process.
 */

const bm25 = require('wink-bm25-text-search');
const embeddingsService = require('./embeddings-service');

const RRF_K = 60;

function tokenize(text) {
  return (text || '')
    .toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .split(/\s+/)
    .filter((t) => t.length > 1);
}

let bm25Engine = null;
let lastEmailIds = null;

function ensureBM25Index(emails) {
  const ids = emails.map((e) => e.id).join(',');
  if (bm25Engine && lastEmailIds === ids) return;
  lastEmailIds = ids;
  bm25Engine = bm25();
  bm25Engine.defineConfig({ fldWeights: { title: 2, body: 1 } });
  bm25Engine.definePrepTasks([tokenize]);
  emails.forEach((e, i) => {
    bm25Engine.addDoc(
      { title: e.subject || '', body: (e.body || '').slice(0, 5000) },
      i
    );
  });
  bm25Engine.consolidate();
}

function sparseSearch(query, limit) {
  if (!bm25Engine) return [];
  const results = bm25Engine.search(query, limit);
  return results;
}

function denseSearch(queryEmbedding, embeddings, emailIds, limit) {
  const scored = emailIds
    .map((id) => {
      const vec = embeddings[id];
      if (!vec) return { id, sim: -1 };
      let dot = 0, normA = 0, normB = 0;
      for (let i = 0; i < vec.length; i++) {
        dot += queryEmbedding[i] * vec[i];
        normA += queryEmbedding[i] * queryEmbedding[i];
        normB += vec[i] * vec[i];
      }
      const sim = normA && normB ? dot / (Math.sqrt(normA) * Math.sqrt(normB)) : 0;
      return { id, sim };
    })
    .filter((s) => s.sim > 0);
  scored.sort((a, b) => b.sim - a.sim);
  return scored.slice(0, limit);
}

function rrfMerge(denseResults, sparseResults, emails) {
  const denseByIdx = {};
  emails.forEach((e, i) => { denseByIdx[e.id] = i; });
  const sparseByIdx = {};
  sparseResults.forEach(([idx, score], rank) => {
    const id = emails[idx] && emails[idx].id;
    if (id) sparseByIdx[id] = rank;
  });
  const scores = {};
  denseResults.forEach((r, rank) => {
    scores[r.id] = (scores[r.id] || 0) + 1 / (RRF_K + rank + 1);
  });
  Object.keys(sparseByIdx).forEach((id) => {
    scores[id] = (scores[id] || 0) + 1 / (RRF_K + sparseByIdx[id] + 1);
  });
  return Object.entries(scores)
    .sort((a, b) => b[1] - a[1])
    .map(([id]) => id);
}

/**
 * Returns email IDs whose embedding similarity to the prompt meets the threshold.
 * Used for prompt-based clusters (no LLM).
 */
async function emailsBySimilarityToPrompt(prompt, embeddings, emailIds, threshold = 0.2) {
  const scored = await emailsBySimilarityToPromptScored(prompt, embeddings, emailIds);
  const above = scored.filter((s) => s.sim >= threshold);
  if (above.length > 0) return above.map((s) => s.id);
  return scored.slice(0, 50).map((s) => s.id);
}

/**
 * Returns all emails with their similarity to the prompt (for user-defined threshold).
 * @returns {Promise<Array<{id: string, sim: number}>>} sorted by sim descending
 */
async function emailsBySimilarityToPromptScored(prompt, embeddings, emailIds) {
  if (!prompt || !embeddings || !emailIds || emailIds.length === 0) return [];
  const queryEmbedding = await embeddingsService.computeQueryEmbedding(prompt);
  const scored = emailIds
    .map((id) => {
      const vec = embeddings[id];
      if (!vec) return { id, sim: 0 };
      let dot = 0, normA = 0, normB = 0;
      for (let i = 0; i < vec.length; i++) {
        dot += queryEmbedding[i] * vec[i];
        normA += queryEmbedding[i] * queryEmbedding[i];
        normB += vec[i] * vec[i];
      }
      const sim = normA && normB ? dot / (Math.sqrt(normA) * Math.sqrt(normB)) : 0;
      return { id, sim };
    });
  scored.sort((a, b) => b.sim - a.sim);
  return scored;
}

function attachSearchScores(email, rank, denseScore, sparseScore) {
  return {
    ...email,
    searchRank: rank,
    denseScore: denseScore != null ? denseScore : null,
    sparseScore: sparseScore != null ? sparseScore : null
  };
}

/** Build sparse score map: email id -> BM25 score (from [idx, score] results). */
function buildSparseScoreById(emails, sparseResults) {
  const out = {};
  sparseResults.forEach((entry) => {
    const idx = Array.isArray(entry) ? entry[0] : entry;
    const score = Array.isArray(entry) && entry.length > 1 ? entry[1] : null;
    const email = emails[idx];
    if (email && score != null) out[email.id] = score;
  });
  return out;
}

/** Ensure every result has both dense and sparse scores when possible. */
async function fillDenseAndSparseForResults(resultEmails, query, emails, embeddings, denseById, sparseScoreById) {
  const byId = {};
  emails.forEach((e) => { byId[e.id] = e; });
  let denseMap = denseById || {};
  let sparseMap = sparseScoreById || {};
  if (embeddings && Object.keys(embeddings).length > 0) {
    try {
      const queryEmbedding = await embeddingsService.computeQueryEmbedding(query);
      const emailIds = emails.map((e) => e.id);
      const denseResults = denseSearch(queryEmbedding, embeddings, emailIds, Math.max(500, resultEmails.length));
      denseResults.forEach((r) => { denseMap[r.id] = r.sim; });
    } catch (_) {}
  }
  if (!sparseScoreById && bm25Engine) {
    const sparseResults = sparseSearch(query, Math.max(500, resultEmails.length));
    sparseMap = buildSparseScoreById(emails, sparseResults);
  }
  return resultEmails.map((e) => ({
    ...e,
    denseScore: e.denseScore != null ? e.denseScore : (denseMap[e.id] ?? null),
    sparseScore: e.sparseScore != null ? e.sparseScore : (sparseMap[e.id] ?? null)
  }));
}

async function hybridSearch(query, emails, embeddings) {
  if (!query || !emails || emails.length === 0) {
    return emails;
  }
  ensureBM25Index(emails);
  const sparseResults = sparseSearch(query, 30);
  const limit = 30;

  const denseById = {};
  const sparseScoreById = {};
  sparseResults.forEach((entry, rank) => {
    const idx = Array.isArray(entry) ? entry[0] : entry;
    const score = Array.isArray(entry) && entry.length > 1 ? entry[1] : null;
    const email = emails[idx];
    if (email) sparseScoreById[email.id] = score;
  });

  let denseResults = [];
  if (embeddings && Object.keys(embeddings).length > 0) {
    try {
      const queryEmbedding = await embeddingsService.computeQueryEmbedding(query);
      const emailIds = emails.map((e) => e.id);
      denseResults = denseSearch(queryEmbedding, embeddings, emailIds, limit);
      denseResults.forEach((r) => { denseById[r.id] = r.sim; });
    } catch (_) {}
  }

  const byId = {};
  emails.forEach((e) => { byId[e.id] = e; });

  if (denseResults.length === 0 && sparseResults.length === 0) {
    const q = query.toLowerCase();
    const filtered = emails.filter(
      (e) =>
        (e.subject && e.subject.toLowerCase().includes(q)) ||
        (e.body && e.body.toLowerCase().includes(q))
    );
    let fallbackResults = filtered.map((e, i) => attachSearchScores(e, i + 1, null, null));
    fallbackResults = await fillDenseAndSparseForResults(fallbackResults, query, emails, embeddings, {}, null);
    return fallbackResults;
  }
  if (denseResults.length === 0) {
    let sparseOnly = sparseResults
      .map((entry, rank) => {
        const idx = Array.isArray(entry) ? entry[0] : entry;
        const score = Array.isArray(entry) && entry.length > 1 ? entry[1] : null;
        const email = emails[idx];
        return email ? attachSearchScores(email, rank + 1, null, score) : null;
      })
      .filter(Boolean);
    sparseOnly = await fillDenseAndSparseForResults(sparseOnly, query, emails, embeddings, {}, sparseScoreById);
    return sparseOnly;
  }
  if (sparseResults.length === 0) {
    let denseOnly = denseResults
      .map((r, rank) => {
        const email = byId[r.id];
        return email ? attachSearchScores(email, rank + 1, r.sim, null) : null;
      })
      .filter(Boolean);
    denseOnly = await fillDenseAndSparseForResults(denseOnly, query, emails, embeddings, denseById, null);
    return denseOnly;
  }

  const mergedIds = rrfMerge(denseResults, sparseResults, emails);
  let merged = mergedIds
    .map((id, rank) => {
      const email = byId[id];
      return email
        ? attachSearchScores(email, rank + 1, denseById[id] ?? null, sparseScoreById[id] ?? null)
        : null;
    })
    .filter(Boolean);
  merged = await fillDenseAndSparseForResults(merged, query, emails, embeddings, denseById, sparseScoreById);
  return merged;
}

module.exports = { hybridSearch, ensureBM25Index, emailsBySimilarityToPrompt, emailsBySimilarityToPromptScored };
