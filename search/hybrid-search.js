/**
 * Hybrid search: delegates to main process (BM25 + dense + RRF).
 */
(function () {
  async function search(query, emails, embeddings, electronAPI) {
    if (!query || !emails || emails.length === 0) return emails;
    if (typeof electronAPI !== 'undefined' && electronAPI.search && electronAPI.search.hybrid) {
      try {
        const res = await electronAPI.search.hybrid(query, emails, embeddings);
        if (res.ok && res.emails) return res.emails;
      } catch (_) {}
    }
    const q = query.toLowerCase();
    return emails.filter(
      (e) =>
        (e.subject && e.subject.toLowerCase().includes(q)) ||
        (e.body && e.body.toLowerCase().includes(q))
    );
  }

  window.HybridSearch = { search };
})();
