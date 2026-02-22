/**
 * Local embeddings service - runs in Electron main process.
 * Uses @huggingface/transformers with Xenova/all-MiniLM-L6-v2 (384-dim).
 * No LLM APIs - fully offline/privacy-first.
 */

/** Strip HTML tags and decode entities for plain text extraction */
function stripHtml(html) {
  if (!html || typeof html !== 'string') return '';
  return html
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .trim();
}

/** Extract text for embedding: subject + body (no images) */
function getTextForEmbedding(email) {
  const subject = (email.subject || '').replace(/\s+/g, ' ').trim();
  let body = email.body || '';
  if (email.bodyIsHtml && body) {
    body = stripHtml(body);
  }
  const text = (subject + ' ' + body).trim().slice(0, 512 * 4); // ~512 tokens
  return text || '(no content)';
}

let pipeline = null;

async function getPipeline() {
  if (pipeline) return pipeline;
  const { pipeline: createPipeline } = await import('@huggingface/transformers');
  pipeline = await createPipeline('feature-extraction', 'Xenova/all-MiniLM-L6-v2');
  return pipeline;
}

/**
 * Compute embeddings for a list of emails.
 * @param {Array<{id:string,subject?:string,body?:string,bodyIsHtml?:boolean}>} emails
 * @param {(progress: {current:number, total:number, message:string}) => void} onProgress
 * @returns {Promise<{[emailId:string]: number[]}>}
 */
async function computeEmbeddings(emails, onProgress) {
  const texts = emails.map((e) => getTextForEmbedding(e));
  const extractor = await getPipeline();

  const BATCH_SIZE = 8;
  const HIDDEN_SIZE = 384;
  const results = {};

  for (let i = 0; i < emails.length; i += BATCH_SIZE) {
    const batch = texts.slice(i, i + BATCH_SIZE);
    const batchEmails = emails.slice(i, i + BATCH_SIZE);

    if (onProgress) {
      onProgress({ current: Math.min(i + BATCH_SIZE, emails.length), total: emails.length, message: 'Embedding...' });
    }

    const output = await extractor(batch, { pooling: 'mean', normalize: true });

    const data = output.data;
    const dims = output.dims || [batch.length, HIDDEN_SIZE];
    const hiddenSize = dims[dims.length - 1] || HIDDEN_SIZE;

    for (let j = 0; j < batchEmails.length; j++) {
      const email = batchEmails[j];
      const start = j * hiddenSize;
      const end = start + hiddenSize;
      const vec = Array.from(data.slice(start, end));
      results[email.id] = vec;
    }
  }

  return results;
}

/**
 * Compute embedding for a single query string (for search).
 * @param {string} query
 * @returns {Promise<number[]>}
 */
async function computeQueryEmbedding(query) {
  const text = (query || '').trim().slice(0, 512) || ' ';
  const extractor = await getPipeline();
  const output = await extractor(text, { pooling: 'mean', normalize: true });
  const data = output.data;
  return Array.from(data);
}

module.exports = {
  getTextForEmbedding,
  computeEmbeddings,
  computeQueryEmbedding,
  stripHtml
};
