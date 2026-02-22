/**
 * Clustering for email categorization. Uses DBSCAN - auto cluster count.
 * Runs in main process.
 */

const DBSCAN = require('density-clustering').DBSCAN;

const CATEGORY_COLORS = ['#EBD98E', '#C89BE6', '#E07ACD', '#7AC8E0', '#8EE07A', '#E0A87A'];

/**
 * Cluster embeddings using DBSCAN.
 * DBSCAN returns clusters as [[idx,idx,...],[idx,...],...] and noise as separate.
 * @param {number[][]} embeddings - Array of embedding vectors
 * @param {string[]} emailIds - Email IDs in same order as embeddings
 * @param {number} eps - Max distance for neighbors (default 0.6 for normalized embeddings)
 * @param {number} minPts - Min points to form cluster (default 2)
 * @returns {{ assignments: {[emailId:string]: string}, meta: {[id:string]: {name:string, color:string}} }}
 */
function cluster(embeddings, emailIds, eps = 0.6, minPts = 2) {
  if (!embeddings || embeddings.length === 0 || !emailIds || emailIds.length === 0) {
    return { assignments: {}, meta: {} };
  }
  const dbscan = new DBSCAN();
  const clusters = dbscan.run(embeddings, eps, minPts);
  const noise = dbscan.noise || [];
  const assignments = {};
  const meta = {};
  clusters.forEach((indices, clusterIdx) => {
    const cid = 'cluster-' + clusterIdx;
    meta[cid] = {
      name: 'Cluster ' + (clusterIdx + 1),
      color: CATEGORY_COLORS[clusterIdx % CATEGORY_COLORS.length]
    };
    indices.forEach((i) => {
      assignments[emailIds[i]] = cid;
    });
  });
  noise.forEach((i) => {
    assignments[emailIds[i]] = 'noise';
  });
  if (noise.length > 0) {
    meta['noise'] = { name: 'Uncategorized', color: '#999' };
  }
  return { assignments, meta };
}

module.exports = { cluster };
