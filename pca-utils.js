/**
 * PCA dimensionality reduction - reduces embeddings to 3D for visualization.
 * Runs in main process; uses ml-pca.
 */

const { PCA } = require('ml-pca');

/**
 * Fit PCA on embeddings and project to 3D.
 * @param {number[][]} embeddings - Array of embedding vectors (same length each)
 * @param {number} nComponents - Number of components (default 3)
 * @returns {{ points: number[][], model: object }} - 3D points and serializable model
 */
function fitAndProject(embeddings, nComponents = 3) {
  if (!embeddings || embeddings.length === 0) {
    return { points: [], model: null };
  }
  const pca = new PCA(embeddings, { nComponents });
  const projected = pca.predict(embeddings, { nComponents });
  const points = projected.to2DArray ? projected.to2DArray() : [];
  if (points.length === 0 && projected.rows) {
    for (let i = 0; i < projected.rows; i++) {
      const row = [];
      for (let j = 0; j < nComponents; j++) row.push(projected.get(i, j));
      points.push(row);
    }
  }
  const jsonModel = pca.toJSON();
  const serializableModel = {
    name: jsonModel.name,
    center: jsonModel.center,
    scale: jsonModel.scale,
    means: jsonModel.means,
    stdevs: jsonModel.stdevs,
    U: jsonModel.U && jsonModel.U.to2DArray ? jsonModel.U.to2DArray() : jsonModel.U,
    S: jsonModel.S,
    excludedFeatures: jsonModel.excludedFeatures || []
  };
  return { points, model: serializableModel };
}

/**
 * Project new embeddings using existing PCA model (e.g. when filtering by category).
 * @param {number[][]} embeddings
 * @param {object} model - Serialized model from fitAndProject
 * @returns {number[][]}
 */
function projectWithModel(embeddings, model) {
  if (!model || !embeddings || embeddings.length === 0) return [];
  const pca = PCA.load(model);
  const projected = pca.predict(embeddings, { nComponents: 3 });
  return projected.to2DArray ? projected.to2DArray() : (() => {
    const out = [];
    for (let i = 0; i < projected.rows; i++) {
      const row = [];
      for (let j = 0; j < 3; j++) row.push(projected.get(i, j));
      out.push(row);
    }
    return out;
  })();
}

module.exports = { fitAndProject, projectWithModel };
