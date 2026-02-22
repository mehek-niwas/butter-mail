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
  sparseResults.forEach(([idx], rank) => {
    const id = emails[idx].id;
    sparseByIdx[id] = rank;
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
async function emailsBySimilarityToPrompt(prompt, embeddings, emailIds, threshold = 0.5) {
  if (!prompt || !embeddings || !emailIds || emailIds.length === 0) return [];
  const queryEmbedding = await embeddingsService.computeQueryEmbedding(prompt);
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
    .filter((s) => s.sim >= threshold);
  scored.sort((a, b) => b.sim - a.sim);
  return scored.map((s) => s.id);
}

async function hybridSearch(query, emails, embeddings) {
  if (!query || !emails || emails.length === 0) {
    return emails;
  }
  ensureBM25Index(emails);
  const sparseResults = sparseSearch(query, 30);
  const limit = 30;

  let denseResults = [];
  if (embeddings && Object.keys(embeddings).length > 0) {
    try {
      const queryEmbedding = await embeddingsService.computeQueryEmbedding(query);
      const emailIds = emails.map((e) => e.id);
      denseResults = denseSearch(queryEmbedding, embeddings, emailIds, limit);
    } catch (_) {}
  }

  if (denseResults.length === 0 && sparseResults.length === 0) {
    const q = query.toLowerCase();
    return emails.filter(
      (e) =>
        (e.subject && e.subject.toLowerCase().includes(q)) ||
        (e.body && e.body.toLowerCase().includes(q))
    );
  }
  if (denseResults.length === 0) {
    const byId = {};
    emails.forEach((e) => { byId[e.id] = e; });
    return sparseResults.map(([idx]) => emails[idx]).filter(Boolean);
  }
  if (sparseResults.length === 0) {
    const byId = {};
    emails.forEach((e) => { byId[e.id] = e; });
    return denseResults.map((r) => byId[r.id]).filter(Boolean);
  }

  const mergedIds = rrfMerge(denseResults, sparseResults, emails);
  const byId = {};
  emails.forEach((e) => { byId[e.id] = e; });
  return mergedIds.map((id) => byId[id]).filter(Boolean);
}

module.exports = { hybridSearch, ensureBM25Index, emailsBySimilarityToPrompt };
