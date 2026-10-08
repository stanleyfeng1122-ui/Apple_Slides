/* Liquid Glass runtime for 1920x1080 slide pages (local Chrome), round 5 (optics unchanged from round 3;
 * round 5 adds the per-box data-fringe override).
 *
 * Optics adapted from rdev/liquid-glass-react (MIT, github.com/rdev/liquid-glass-react @ ac48eab):
 * an SVG feDisplacementMap driven by a displacement map, three per-channel scales (chromatic
 * fringe) recombined with screen blends, light blur + saturation, hollow specular rim, no fill.
 *
 * The map is a rounded-rect lens generated per box:
 *  - rim band (width = bezel): samples OUTWARD along the edge normal plus a radial term, so content
 *    just outside the pane is pulled into the rim and compressed; lines crossing a corner curl;
 *  - centre: a mild inward (magnifying) term, like a thick convex pane.
 * The lens filters a clone of the page .backdrop, because backdrop-filter cannot sample outside
 * the box. Everything that may sit under glass lives in the page's .backdrop.
 *
 * Authoring: <div class="glass" data-role="glass" data-tone="dark|light" data-tint="0.2"
 *   style="left;top;width;height;border-radius"> ...native data-role="text" children... </div>
 * Optional per box: data-bezel (px), data-pull (max rim pull in px), data-blur (px), data-mag,
 * data-radial (weight of the radial term in the rim direction; lines that meet a rim at a right angle
 * kink only through this term), data-fringe (channel split in px per step; default GLASS_FRINGE 1.0; use
 * 0.6 over high-frequency texture such as a watch mesh, where 1.0 sparkles along the rim).
 * Backdrop bleed: an element in .backdrop with data-lens-src + data-lens-box="left,top,width,height" is
 * swapped for that larger source inside the lens only, so rims at the edge of a photo sample more photo
 * instead of the page colour (the page itself keeps the clipped original).
 * data-tint is written by calibrate.py from the measured content under each box's text.
 */
(function () {
  const NS = 'http://www.w3.org/2000/svg';
  const FRINGE = +(window.GLASS_FRINGE ?? 1.0);   // px channel split per step at full displacement: R/G/B = 0/1/2 px
  const MAG = +(window.GLASS_MAG ?? 0.06);   // centre magnification: content under the pane looks ~6% larger
  const RADIAL = +(window.GLASS_RADIAL ?? 0.8);   // radial share of the rim pull direction (round 2: 0.45)
  const clamp = (v, a, b) => Math.min(Math.max(v, a), b);

  // Displacement in px: centre magnifies by MAG (samples inward), the rim band samples OUTWARD by up
  // to `pull` px along the edge normal plus a radial term. Returns two maps: the lens (normalised by
  // its max |px|) and the rim-only direction used for the colour fringe, so fringe stays on the rim.
  // The maps extend M px beyond the box with the rim value continued outward (round 3). Before, the area
  // outside the box had no map, so the 1-2 px channel shift at the very edge read garbage and drew a
  // thin blue line along bottom rims.
  function lensMaps(w, h, r, bezel, pull, mag, radial, M) {
    const W = w + 2 * M, H = h + 2 * M;
    const disp = new Float32Array(W * H * 4);
    let max = 1e-6;
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const px = x + 0.5 - W / 2, py = y + 0.5 - H / 2;
      const qx = Math.abs(px) - (w / 2 - r), qy = Math.abs(py) - (h / 2 - r);
      let depth, nx, ny;
      if (qx > 0 && qy > 0) {                        // corner arc: normal rotates -> curls
        const l = Math.hypot(qx, qy) || 1e-6;
        depth = r - l; nx = qx / l * Math.sign(px); ny = qy / l * Math.sign(py);
      } else if (qx > qy) { depth = r - qx; nx = Math.sign(px); ny = 0; }
      else { depth = r - qy; nx = 0; ny = Math.sign(py); }
      const m = Math.pow(1 - clamp(depth / bezel, 0, 1), 2);          // 1 at the edge and outside, 0 inside
      let dx = 0.7 * nx + radial * px / (w / 2), dy = 0.7 * ny + radial * py / (h / 2);
      const len = Math.hypot(dx, dy) || 1; dx /= len; dy /= len;         // outward rim direction
      const i = 4 * (y * W + x);
      disp[i] = m * pull * dx - (1 - m) * mag * px; disp[i + 1] = m * pull * dy - (1 - m) * mag * py;
      disp[i + 2] = m * dx; disp[i + 3] = m * dy;
      max = Math.max(max, Math.abs(disp[i]), Math.abs(disp[i + 1]));
    }
    const toMap = (k, norm) => {
      const c = document.createElement('canvas');
      c.width = W; c.height = H;
      const ctx = c.getContext('2d'), img = ctx.createImageData(W, H), d = img.data;
      for (let i = 0; i < W * H; i++) {
        d[4 * i] = 128 + disp[4 * i + k] / norm * 127; d[4 * i + 1] = 128 + disp[4 * i + k + 1] / norm * 127;
        d[4 * i + 2] = 128; d[4 * i + 3] = 255;
      }
      ctx.putImageData(img, 0, 0);
      return c.toDataURL('image/png');
    };
    return {lens: toMap(0, max), rim: toMap(2, 1), max};
  }

  function el(tag, attrs) {
    const e = document.createElementNS(NS, tag);
    for (const k in attrs) e.setAttribute(k, attrs[k]);
    return e;
  }

  // feDisplacementMap: out(x) = src(x + scale*(map-0.5)), so the max shift in px is scale/2.
  // Stage 1 bends the backdrop with the lens map; stage 2 shifts G and B a further 1/2 px outward
  // along the rim only (channel split -> colour fringe on edges that cross the rim).
  function filterFor(id, box, bezel, pull, mag, radial, fringe) {
    const M = Math.ceil(2 * fringe) + 3;
    const maps = lensMaps(box.w, box.h, box.r, bezel, pull, mag, radial, M);
    const pad = Math.ceil(maps.max + 3 * fringe + 40);
    const f = el('filter', {id, filterUnits: 'userSpaceOnUse', primitiveUnits: 'userSpaceOnUse',
      x: box.x - pad, y: box.y - pad, width: box.w + 2 * pad, height: box.h + 2 * pad,
      'color-interpolation-filters': 'sRGB'});
    const img = (href, result) => el('feImage', {href, x: box.x - M, y: box.y - M, width: box.w + 2 * M, height: box.h + 2 * M,
      preserveAspectRatio: 'none', result});
    f.append(img(maps.lens, 'lens'), img(maps.rim, 'rim'));
    f.append(el('feDisplacementMap', {in: 'SourceGraphic', in2: 'lens', scale: 2 * maps.max,
      xChannelSelector: 'R', yChannelSelector: 'G', result: 'bent'}));
    const keep = ['1 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 1 0',
                  '0 0 0 0 0  0 1 0 0 0  0 0 0 0 0  0 0 0 1 0',
                  '0 0 0 0 0  0 0 0 0 0  0 0 1 0 0  0 0 0 1 0'];
    [0, 1, 2].forEach(i => {
      f.append(el('feDisplacementMap', {in: 'bent', in2: 'rim', scale: 2 * fringe * i,
        xChannelSelector: 'R', yChannelSelector: 'G', result: 'd' + i}));
      f.append(el('feColorMatrix', {in: 'd' + i, type: 'matrix', values: keep[i], result: 'c' + i}));
    });
    f.append(el('feBlend', {in: 'c0', in2: 'c1', mode: 'screen', result: 'c01'}));
    f.append(el('feBlend', {in: 'c01', in2: 'c2', mode: 'screen'}));
    return f;
  }

  function build() {
    const defs = el('svg', {width: 0, height: 0, 'aria-hidden': 'true'});
    defs.style.position = 'absolute';
    document.body.prepend(defs);
    document.querySelectorAll('.glass').forEach((g, n) => {
      const page = g.closest('.page'), cs = getComputedStyle(g), ps = getComputedStyle(page);
      const box = {x: parseFloat(cs.left), y: parseFloat(cs.top), w: Math.round(parseFloat(cs.width)),
                   h: Math.round(parseFloat(cs.height)), r: parseFloat(cs.borderTopLeftRadius) || 0};
      box.r = Math.min(box.r, box.w / 2, box.h / 2);
      const tone = g.dataset.tone || ((page.dataset.theme === 'white' || page.classList.contains('white')) ? 'light' : 'dark');
      g.classList.add(tone);
      // Round 3: one bezel rule for capsules and tiles, and no light-tone pull factor (it hid the lens on white).
      const bezel = +(g.dataset.bezel || clamp(0.3 * Math.min(box.w, box.h), 16, 56));
      const pull = +(g.dataset.pull || clamp(0.35 * bezel + 6, 10, 26));
      const radial = +(g.dataset.radial || RADIAL);
      if (g.dataset.tint) g.style.setProperty('--a', g.dataset.tint);
      if (g.dataset.blur) g.style.setProperty('--glass-blur', g.dataset.blur + 'px');
      const id = 'lg' + n;
      const fringe = +(g.dataset.fringe || FRINGE);
      defs.append(filterFor(id, box, bezel, pull, +(g.dataset.mag || MAG), radial, fringe));
      const optic = document.createElement('div');
      optic.className = 'optic';
      optic.style.cssText = `left:${-box.x}px;top:${-box.y}px;background:${ps.backgroundColor};` +
        `filter:blur(var(--glass-blur)) url(#${id}) saturate(var(--glass-saturate))`;
      const clone = page.querySelector('.backdrop').cloneNode(true);
      // Same markup and CSS as the real backdrop; converters skip anything inside .optic.
      clone.querySelectorAll('[id]').forEach(e => e.removeAttribute('id'));
      clone.querySelectorAll('[data-lens-src]').forEach(e => {
        const [l, t, w, h] = e.dataset.lensBox.split(',').map(Number);
        e.src = e.dataset.lensSrc;
        Object.assign(e.style, {left: l + 'px', top: t + 'px', width: w + 'px', height: h + 'px'});
      });
      clone.setAttribute('aria-hidden', 'true');
      // Light glass: the rim also refracts the soft shadow the pane casts on the page just outside it, so
      // even over plain white the rim reads as a lens (darker toward the bottom, where the shadow falls).
      if (tone === 'light') {
        const sh = document.createElement('div');
        sh.style.cssText = `position:absolute;left:${box.x}px;top:${box.y}px;width:${box.w}px;height:${box.h}px;` +
          `border-radius:${box.r}px;box-shadow:${cs.boxShadow}`;
        clone.append(sh);
      }
      optic.append(clone);
      const lens = document.createElement('div'); lens.className = 'lens'; lens.append(optic);
      const tint = document.createElement('div'); tint.className = 'tint';
      const rim = document.createElement('div'); rim.className = 'rim';
      g.prepend(lens, tint, rim);
      g.dataset.optics = JSON.stringify({tone, bezel: +bezel.toFixed(1), pull: +pull.toFixed(1), radial, fringe});
    });
  }
  window.buildGlass = build;
  window.glassReady = false;
  if (!window.GLASS_DEFER_BUILD) { build(); window.glassReady = true; }
})();
