/**
 * Clustering for email categorization. Uses DBSCAN - auto cluster count.
 * Runs in main process.
 */

const DBSCAN = require('density-clustering').DBSCAN;

const CATEGORY_COLORS = ['#B8952E', '#7B4BA6', '#A8348A', '#2A7B8A', '#4A9B3A', '#9B5A2A'];
const MIN_AUTO_CLUSTER_EMAILS = 10;

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
  const noiseSet = new Set(dbscan.noise || []);
  const keptClusters = [];
  clusters.forEach((indices) => {
    if (indices.length >= MIN_AUTO_CLUSTER_EMAILS) {
      keptClusters.push(indices);
    } else {
      indices.forEach((i) => noiseSet.add(i));
    }
  });
  const noise = Array.from(noiseSet);
  const assignments = {};
  const meta = {};
  keptClusters.forEach((indices, clusterIdx) => {
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
