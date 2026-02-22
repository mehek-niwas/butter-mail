/**
 * 3D PCA Graph view - Three.js scatter plot with axes and modern startup aesthetic.
 */
(function () {
  let scene, camera, renderer, controls, pointsGroup, axesGroup, raycaster, mouse;

  function createSharpPointTexture() {
    const size = 64;
    const canvas = document.createElement('canvas');
    canvas.width = size;
    canvas.height = size;
    const ctx = canvas.getContext('2d');
    const r = size / 2 - 1;
    ctx.fillStyle = 'white';
    ctx.beginPath();
    ctx.arc(size / 2, size / 2, r, 0, Math.PI * 2);
    ctx.fill();
    const tex = new THREE.CanvasTexture(canvas);
    tex.needsUpdate = true;
    tex.minFilter = THREE.NearestFilter;
    tex.magFilter = THREE.NearestFilter;
    return tex;
  }

  const AXIS_COLOR = 0x1a365d;

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
    scene.background = new THREE.Color(0xf5f2e8);

    camera = new THREE.PerspectiveCamera(55, width / height, 0.1, 1000);
    camera.position.set(6, 5, 6);

    renderer = new THREE.WebGLRenderer({ canvas, antialias: true, alpha: true });
    renderer.setSize(width, height);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    renderer.setClearColor(0xf5f2e8, 1);

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

    const positions = [];
    const colors = [];
    arr.forEach(([emailId, p]) => {
      positions.push((p[0] || 0) * scale, (p[1] || 0) * scale, (p[2] || 0) * scale);
      const email = emailsById && emailsById[emailId];
      const catId = email && email.categoryId;
      const hex = (catId && window.getCategoryColor) ? window.getCategoryColor(catId) : '#B8952E';
      const c = new THREE.Color(hex);
      colors.push(c.r, c.g, c.b);
    });

    const geometry = new THREE.BufferGeometry();
    geometry.setAttribute('position', new THREE.Float32BufferAttribute(positions, 3));
    geometry.setAttribute('color', new THREE.Float32BufferAttribute(colors, 3));

    const pointTexture = createSharpPointTexture();
    const material = new THREE.PointsMaterial({
      size: 0.32,
      vertexColors: true,
      sizeAttenuation: true,
      map: pointTexture,
      transparent: false,
      opacity: 1,
      depthWrite: true,
      blending: THREE.NormalBlending
    });
    const points = new THREE.Points(geometry, material);
    points.userData.entries = arr;
    points.userData.emailsById = emailsById || {};
    pointsGroup.add(points);
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
