/**
 * 3D PCA Graph view - Three.js scatter plot with axes and modern startup aesthetic.
 */
(function () {
  let scene, camera, renderer, controls, pointsGroup, axesGroup, raycaster, mouse;

  const AXIS_COLOR = 0x4a3a1c;
  const BG_COLOR = 0x0a0805;
  const DEFAULT_POINT_COLOR = '#ffcc4d';

  const textureCache = {};

  function makeSolidDotTexture(hex) {
    const size = 64;
    const canvas = document.createElement('canvas');
    canvas.width = size;
    canvas.height = size;
    const ctx = canvas.getContext('2d');
    const cx = size / 2;
    const cy = size / 2;
    const r = size / 2 - 2;
    const grad = ctx.createRadialGradient(cx, cy, 0, cx, cy, r);
    grad.addColorStop(0, hex);
    grad.addColorStop(0.7, hex);
    grad.addColorStop(1, hex + '00');
    ctx.fillStyle = grad;
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.fill();
    const tex = new THREE.CanvasTexture(canvas);
    tex.needsUpdate = true;
    return tex;
  }

  function makeSwirlTexture(colors) {
    const size = 96;
    const canvas = document.createElement('canvas');
    canvas.width = size;
    canvas.height = size;
    const ctx = canvas.getContext('2d');
    const cx = size / 2;
    const cy = size / 2;
    const r = size / 2 - 2;
    const n = colors.length;
    const sweep = (Math.PI * 2) / n;
    const layers = 3;
    for (let layer = 0; layer < layers; layer++) {
      const inner = (r * layer) / layers;
      const outer = (r * (layer + 1)) / layers;
      const rotation = (layer * Math.PI) / (n * 2);
      for (let i = 0; i < n; i++) {
        const start = i * sweep + rotation;
        const end = start + sweep + 0.04;
        ctx.beginPath();
        ctx.moveTo(cx + Math.cos(start) * inner, cy + Math.sin(start) * inner);
        ctx.arc(cx, cy, outer, start, end, false);
        ctx.arc(cx, cy, inner, end, start, true);
        ctx.closePath();
        ctx.fillStyle = colors[i];
        ctx.globalAlpha = 0.85;
        ctx.fill();
      }
    }
    ctx.globalAlpha = 1;
    const fade = ctx.createRadialGradient(cx, cy, r * 0.65, cx, cy, r);
    fade.addColorStop(0, 'rgba(0,0,0,0)');
    fade.addColorStop(1, 'rgba(0,0,0,1)');
    ctx.globalCompositeOperation = 'destination-out';
    ctx.fillStyle = fade;
    ctx.fillRect(0, 0, size, size);
    ctx.globalCompositeOperation = 'source-over';
    const tex = new THREE.CanvasTexture(canvas);
    tex.needsUpdate = true;
    return tex;
  }

  function getGroupTexture(colors) {
    const key = colors.length === 0 ? '__default__' : colors.slice().sort().join('|');
    if (textureCache[key]) return textureCache[key];
    let tex;
    if (colors.length === 0) tex = makeSolidDotTexture(DEFAULT_POINT_COLOR);
    else if (colors.length === 1) tex = makeSolidDotTexture(colors[0]);
    else tex = makeSwirlTexture(colors);
    textureCache[key] = tex;
    return tex;
  }

  function colorSetKey(colors) {
    if (!colors || !colors.length) return '__default__';
    return colors.slice().sort().join('|');
  }

  function createAxes(extent) {
    const group = new THREE.Group();
    const axisLen = extent || 4;

    const xGeom = new THREE.BufferGeometry().setFromPoints([
      new THREE.Vector3(-axisLen, 0, 0),
      new THREE.Vector3(axisLen, 0, 0)
    ]);
    const yGeom = new THREE.BufferGeometry().setFromPoints([
      new THREE.Vector3(0, -axisLen, 0),
      new THREE.Vector3(0, axisLen, 0)
    ]);
    const zGeom = new THREE.BufferGeometry().setFromPoints([
      new THREE.Vector3(0, 0, -axisLen),
      new THREE.Vector3(0, 0, axisLen)
    ]);

    const xAxis = new THREE.Line(xGeom, new THREE.LineBasicMaterial({ color: AXIS_COLOR }));
    const yAxis = new THREE.Line(yGeom, new THREE.LineBasicMaterial({ color: AXIS_COLOR }));
    const zAxis = new THREE.Line(zGeom, new THREE.LineBasicMaterial({ color: AXIS_COLOR }));
    xAxis.name = 'xAxis';
    yAxis.name = 'yAxis';
    zAxis.name = 'zAxis';

    group.add(xAxis);
    group.add(yAxis);
    group.add(zAxis);

    return group;
  }

  function init(containerId) {
    const container = document.getElementById(containerId || 'graph-container');
    if (!container) return;
    const canvas = document.getElementById('graph-canvas');
    if (!canvas) return;

    const width = container.clientWidth;
    const height = container.clientHeight;

    scene = new THREE.Scene();
    scene.background = new THREE.Color(BG_COLOR);

    camera = new THREE.PerspectiveCamera(55, width / height, 0.1, 1000);
    camera.position.set(6, 5, 6);

    renderer = new THREE.WebGLRenderer({ canvas, antialias: true, alpha: true });
    renderer.setSize(width, height);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    renderer.setClearColor(BG_COLOR, 1);

    if (typeof THREE.OrbitControls !== 'undefined') {
      controls = new THREE.OrbitControls(camera, renderer.domElement);
      controls.enableDamping = true;
      controls.dampingFactor = 0.06;
      controls.minDistance = 2;
      controls.maxDistance = 25;
    }

    raycaster = new THREE.Raycaster();
    mouse = new THREE.Vector2();

    axesGroup = createAxes(4);
    scene.add(axesGroup);

    pointsGroup = new THREE.Group();
    scene.add(pointsGroup);

    window.addEventListener('resize', onResize);
    canvas.addEventListener('click', onCanvasClick);
    canvas.addEventListener('mousemove', onMouseMove);
  }

  function onResize() {
    const container = document.getElementById('graph-container');
    if (!container) return;
    const width = container.clientWidth;
    const height = container.clientHeight;
    camera.aspect = width / height;
    camera.updateProjectionMatrix();
    renderer.setSize(width, height);
  }

  function getMouseNDC(event) {
    const rect = event.target.getBoundingClientRect();
    return {
      x: ((event.clientX - rect.left) / rect.width) * 2 - 1,
      y: -((event.clientY - rect.top) / rect.height) * 2 + 1
    };
  }

  function onCanvasClick(event) {
    if (!raycaster || !pointsGroup || pointsGroup.children.length === 0) return;
    const ndc = getMouseNDC(event);
    mouse.x = ndc.x;
    mouse.y = ndc.y;
    raycaster.setFromCamera(mouse, camera);
    raycaster.params.Points = raycaster.params.Points || {};
    raycaster.params.Points.threshold = 0.35;
    const intersects = raycaster.intersectObjects(pointsGroup.children, true);
    if (intersects.length > 0) {
      const obj = intersects[0].object;
      const idx = intersects[0].index;
      const entries = obj.userData && obj.userData.entries;
      if (entries && idx >= 0 && idx < entries.length) {
        const emailId = entries[idx][0];
        if (window.onGraphPointClick) window.onGraphPointClick(emailId);
      }
    }
  }

  function onMouseMove(event) {
    const tooltip = document.getElementById('graph-tooltip');
    const coordsEl = document.getElementById('graph-coords');
    if (!raycaster || !pointsGroup) return;
    const ndc = getMouseNDC(event);
    mouse.x = ndc.x;
    mouse.y = ndc.y;
    raycaster.setFromCamera(mouse, camera);
    raycaster.params.Points = raycaster.params.Points || {};
    raycaster.params.Points.threshold = 0.35;
    const intersects = raycaster.intersectObjects(pointsGroup.children, true);
    if (intersects.length > 0) {
      const obj = intersects[0].object;
      const idx = intersects[0].index;
      const entries = obj.userData && obj.userData.entries;
      if (entries && idx >= 0 && idx < entries.length) {
        const emailId = entries[idx][0];
        const p = obj.userData.entries[idx][1];
        const x = (p[0] || 0).toFixed(2);
        const y = (p[1] || 0).toFixed(2);
        const z = (p[2] || 0).toFixed(2);
        if (coordsEl) coordsEl.textContent = 'X: ' + x + '  Y: ' + y + '  Z: ' + z;
        const subject = obj.userData.emailsById && obj.userData.emailsById[emailId] ? obj.userData.emailsById[emailId].subject : '';
        if (tooltip) {
          tooltip.textContent = (subject || '(no subject)') + ' [' + x + ', ' + y + ', ' + z + ']';
          tooltip.classList.remove('hidden');
          const rect = event.target.getBoundingClientRect();
          tooltip.style.left = (event.clientX - rect.left + 10) + 'px';
          tooltip.style.top = (event.clientY - rect.top + 10) + 'px';
        }
      }
    } else {
      if (coordsEl) coordsEl.textContent = 'X: —  Y: —  Z: —';
      if (tooltip) tooltip.classList.add('hidden');
    }
  }

  function setDataSimple(pointsData, emailsById) {
    if (!pointsGroup) return;
    while (pointsGroup.children.length > 0) {
      pointsGroup.remove(pointsGroup.children[0]);
    }
    if (!pointsData || Object.keys(pointsData).length === 0) return;

    const arr = Object.entries(pointsData);
    let maxAbs = 1;
    arr.forEach(([, p]) => {
      maxAbs = Math.max(maxAbs, Math.abs(p[0] || 0), Math.abs(p[1] || 0), Math.abs(p[2] || 0));
    });
    const scale = 3.2 / maxAbs;

    const groups = {};
    arr.forEach(([emailId, p]) => {
      const email = emailsById && emailsById[emailId];
      const colors = (email && Array.isArray(email.clusterColors)) ? email.clusterColors : [];
      const key = colorSetKey(colors);
      if (!groups[key]) groups[key] = { colors, entries: [], positions: [] };
      groups[key].entries.push([emailId, p]);
      groups[key].positions.push((p[0] || 0) * scale, (p[1] || 0) * scale, (p[2] || 0) * scale);
    });

    Object.keys(groups).forEach((key) => {
      const grp = groups[key];
      const geometry = new THREE.BufferGeometry();
      geometry.setAttribute('position', new THREE.Float32BufferAttribute(grp.positions, 3));
      const texture = getGroupTexture(grp.colors);
      const material = new THREE.PointsMaterial({
        size: grp.colors.length > 1 ? 0.42 : 0.32,
        sizeAttenuation: true,
        map: texture,
        transparent: true,
        alphaTest: 0.05,
        depthWrite: false,
        blending: THREE.NormalBlending
      });
      const points = new THREE.Points(geometry, material);
      points.userData.entries = grp.entries;
      points.userData.emailsById = emailsById || {};
      pointsGroup.add(points);
    });
  }

  function render(pointsData, emailsById) {
    setDataSimple(pointsData, emailsById);
  }

  function animate() {
    if (!renderer || !scene || !camera) return;
    requestAnimationFrame(animate);
    if (controls) controls.update();
    renderer.render(scene, camera);
  }

  window.GraphView = {
    init,
    render,
    setData: setDataSimple,
    animate
  };
})();
