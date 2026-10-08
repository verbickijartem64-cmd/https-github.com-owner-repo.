// Векторные иллюстрации для версии 6 (SVG → PNG через sharp). Без текста внутри: подписи делаются в PowerPoint.
const sharp = require('sharp');

// детерминированный генератор (одинаковая картинка при каждой сборке)
function rng(seed) { let s = seed >>> 0; return () => ((s = (s * 1664525 + 1013904223) >>> 0) / 4294967296); }

function speckles(r, n, x0, x1, y0, y1, colors, rmin, rmax, op) {
  let out = '';
  for (let i = 0; i < n; i++) {
    const x = x0 + r() * (x1 - x0), y = y0 + r() * (y1 - y0), rad = rmin + r() * (rmax - rmin);
    const c = colors[Math.floor(r() * colors.length)];
    out += `<ellipse cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" rx="${rad.toFixed(1)}" ry="${(rad * (0.6 + r() * 0.4)).toFixed(1)}" fill="${c}" opacity="${op}"/>`;
  }
  return out;
}

// 1. Титул: почвенный керн с датчиком (кадр 1000 × 1433, справа и снизу «уходит» за край)
function coverSVG() {
  const r = rng(616);
  const W = 1000, H = 1433, X0 = 150, TOP = 330, AB = 700, BC = 1080;
  const wav = (y, a, ph) => `M${X0} ${y} C ${X0 + 220} ${y - a + ph}, ${X0 + 430} ${y + a}, ${X0 + 600} ${y - ph} S ${W - 40} ${y + a * 0.6}, ${W} ${y}`;
  let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
  <defs>
    <linearGradient id="gA" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#D2AC7C"/><stop offset="1" stop-color="#BA8F60"/></linearGradient>
    <linearGradient id="gB" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#A4704A"/><stop offset="1" stop-color="#875737"/></linearGradient>
    <linearGradient id="gC" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#7A432C"/><stop offset="1" stop-color="#55291B"/></linearGradient>
    <linearGradient id="rod" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#C9CCCD"/><stop offset="0.45" stop-color="#F4F5F5"/><stop offset="1" stop-color="#A9AEB0"/></linearGradient>
    <linearGradient id="head" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#FFFFFF"/><stop offset="1" stop-color="#DCD8D2"/></linearGradient>
    <linearGradient id="edge" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#000" stop-opacity="0.35"/><stop offset="0.12" stop-color="#000" stop-opacity="0"/></linearGradient>
    <filter id="glow" x="-50%" y="-10%" width="200%" height="120%"><feGaussianBlur stdDeviation="9"/></filter>
    <clipPath id="core"><rect x="${X0}" y="${TOP}" width="${W - X0 + 60}" height="${H - TOP + 60}" rx="38"/></clipPath>
  </defs>
  <g clip-path="url(#core)">
    <rect x="0" y="0" width="${W}" height="${H}" fill="url(#gC)"/>
    <path d="${wav(BC, 26, 8)} L ${W} 0 L ${X0} 0 Z" fill="url(#gB)"/>
    <path d="${wav(AB, 22, -6)} L ${W} 0 L ${X0} 0 Z" fill="url(#gA)"/>
    ${speckles(r, 260, X0, W, TOP + 10, AB - 10, ['#E6CBA3', '#9C7448', '#F0DDBF'], 2, 7, 0.55)}
    ${speckles(r, 240, X0, W, AB + 15, BC - 10, ['#C08A5F', '#6E4127', '#B98258'], 2, 6, 0.5)}
    ${speckles(r, 260, X0, W, BC + 15, H, ['#8E5136', '#3F1D12', '#9A5C3E'], 1.5, 5, 0.5)}
    <rect x="0" y="0" width="${W}" height="${H}" fill="url(#edge)"/>
  </g>`;
  // трава / виноградник по краю керна
  for (let i = 0; i < 26; i++) {
    const x = X0 + 30 + i * 32 + r() * 10, h = 22 + r() * 34, lean = (r() - 0.5) * 26;
    s += `<path d="M${x.toFixed(1)} ${TOP + 2} q ${(lean / 2).toFixed(1)} ${(-h / 2).toFixed(1)} ${lean.toFixed(1)} ${(-h).toFixed(1)}" stroke="#7E9B63" stroke-width="4" stroke-linecap="round" fill="none" opacity="0.9"/>`;
  }
  // кабель, корпус и штанги датчика
  const cx = 590, R1 = 556, R2 = 624, TIP = 1170;
  s += `<path d="M${cx} 205 C ${cx} 120, ${cx + 120} 90, ${cx + 170} 0" stroke="#2A2420" stroke-width="16" fill="none" stroke-linecap="round"/>`;
  s += `<path d="M${cx} 205 C ${cx} 120, ${cx + 120} 90, ${cx + 170} 0" stroke="#4A403A" stroke-width="8" fill="none" stroke-linecap="round"/>`;
  [R1, R2].forEach(x => { s += `<rect x="${x - 9}" y="${TOP - 20}" width="18" height="${TIP - TOP + 20}" rx="9" fill="url(#rod)"/>`; });
  s += `<rect x="${cx - 78}" y="200" width="156" height="${TOP - 190}" rx="18" fill="url(#head)"/>`;
  s += `<rect x="${cx - 52}" y="236" width="104" height="10" rx="5" fill="#008C9E"/>`;
  // импульс: вниз вдоль первой штанги и обратно вдоль второй
  const down = `M${R1 - 34} ${TOP + 40} L ${R1 - 34} ${TIP - 30}`, up = `M${R2 + 34} ${TIP - 30} L ${R2 + 34} ${TOP + 40}`;
  [[down, 1], [up, 0.75]].forEach(([d, op]) => {
    s += `<path d="${d}" stroke="#46D3E3" stroke-width="16" opacity="${0.55 * op}" filter="url(#glow)" fill="none"/>`;
    s += `<path d="${d}" stroke="#7FE4EF" stroke-width="5" stroke-dasharray="26 16" opacity="${op}" fill="none" stroke-linecap="round"/>`;
  });
  const arrow = (x, y, dir) => `<path d="M${x - 16} ${y - 18 * dir} L ${x} ${y} L ${x + 16} ${y - 18 * dir}" stroke="#7FE4EF" stroke-width="5" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`;
  s += arrow(R1 - 34, TIP - 26, 1) + arrow(R2 + 34, TOP + 36, -1);
  s += `</svg>`;
  return s;
}

// 2. Датчик в почве и «лупа» с тремя фазами: минералы, вода, воздух (кадр 1100 × 620)
// Возвращает SVG; координаты опорных точек для подписей — в SENSOR_ANCHORS (в пикселях кадра).
const SENSOR = { W:1100, H:620, mag:{ cx:868, cy:372, r:212 }, focus:{ cx:520, cy:430, r:44 } };
const GRAINS = [ // зёрна в системе координат лупы (центр 0,0): x, y, rx, ry, поворот, цвет
  [-112, -112, 74, 56, 18, '#D9BC92'], [36, -150, 70, 50, -12, '#C9A578'], [150, -52, 58, 70, 28, '#E1C9A1'],
  [-168, 22, 56, 66, -18, '#C9A578'], [-6, -8, 66, 50, 14, '#D2B184'], [118, 98, 72, 52, -24, '#D9BC92'],
  [-96, 142, 68, 48, 8, '#E1C9A1'], [36, 208, 62, 40, 0, '#C9A578'], [-218, -150, 60, 50, 0, '#D2B184'],
  [222, 140, 50, 60, 0, '#C9A578'], [190, -190, 60, 46, 20, '#D2B184'] ];
const BRIDGES = [[-50, -60, 34, 22, 30], [80, -95, 30, 20, -40], [62, 46, 30, 22, 40], [-88, 70, 30, 24, -30], [-36, 182, 28, 18, 0], [176, 30, 24, 20, 70]];
const SENSOR_ANCHORS = { mineral:[868 + 36, 372 - 150], water:[868 + 62, 372 + 46], air:[868 - 10, 372 + 80] };
function sensorSVG() {
  const r = rng(42);
  const { W, H, mag, focus } = SENSOR, BX = 20, BY = 228, BW = 640, BH = 372;
  let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
  <defs>
    <linearGradient id="soil" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#BE9568"/><stop offset="1" stop-color="#8E623F"/></linearGradient>
    <linearGradient id="rod" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#B9BEC0"/><stop offset="0.45" stop-color="#F4F5F5"/><stop offset="1" stop-color="#9DA3A5"/></linearGradient>
    <radialGradient id="air" cx="0.5" cy="0.45" r="0.6"><stop offset="0" stop-color="#FBF8F3"/><stop offset="1" stop-color="#EEE7DC"/></radialGradient>
    <filter id="glow" x="-50%" y="-10%" width="200%" height="120%"><feGaussianBlur stdDeviation="6"/></filter>
    <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%"><feGaussianBlur stdDeviation="10"/></filter>
    <clipPath id="blk"><rect x="${BX}" y="${BY}" width="${BW}" height="${BH}" rx="26"/></clipPath>
    <clipPath id="lens"><circle cx="${mag.cx}" cy="${mag.cy}" r="${mag.r}"/></clipPath>
  </defs>
  <g clip-path="url(#blk)">
    <rect x="${BX}" y="${BY}" width="${BW}" height="${BH}" fill="url(#soil)"/>
    ${speckles(r, 420, BX, BX + BW, BY + 6, BY + BH, ['#DDBF95', '#7D5434', '#E8D2AE', '#A57A51'], 1.8, 5, 0.65)}
  </g>`;
  // датчик: штанги, корпус, кабель, импульс
  const R1 = 180, R2 = 262, TIP = 545;
  [R1, R2].forEach(x => { s += `<rect x="${x - 8}" y="${BY - 14}" width="16" height="${TIP - BY + 14}" rx="8" fill="url(#rod)"/>`; });
  s += `<path d="M222 98 C 222 52, 300 30, 372 0" stroke="#2E2622" stroke-width="12" fill="none" stroke-linecap="round"/>`;
  s += `<rect x="120" y="96" width="204" height="${BY - 82}" rx="18" fill="#2E2622"/>`;
  s += `<rect x="146" y="128" width="152" height="10" rx="5" fill="#008C9E"/>`;
  const down = `M${R1 - 34} ${BY + 26} L ${R1 - 34} ${TIP - 18}`, up = `M${R2 + 34} ${TIP - 18} L ${R2 + 34} ${BY + 26}`;
  [[down, 1], [up, 0.8]].forEach(([d, op]) => {
    s += `<path d="${d}" stroke="#7FE4EF" stroke-width="13" opacity="${0.45 * op}" filter="url(#glow)" fill="none"/>`;
    s += `<path d="${d}" stroke="#DDF7FA" stroke-width="5" stroke-dasharray="20 12" opacity="${op}" fill="none" stroke-linecap="round"/>`;
  });
  const arrow = (x, y, dir) => `<path d="M${x - 13} ${y - 15 * dir} L ${x} ${y} L ${x + 13} ${y - 15 * dir}" stroke="#DDF7FA" stroke-width="5" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`;
  s += arrow(R1 - 34, TIP - 14, 1) + arrow(R2 + 34, BY + 22, -1);
  // лупа: касательные от малого круга к большому
  const dx = mag.cx - focus.cx, dy = mag.cy - focus.cy, D = Math.hypot(dx, dy), th = Math.atan2(dy, dx), ph = Math.acos((focus.r - mag.r) / D);
  [th + ph, th - ph].forEach(a => {
    const x1 = focus.cx + focus.r * Math.cos(a), y1 = focus.cy + focus.r * Math.sin(a), x2 = mag.cx + mag.r * Math.cos(a), y2 = mag.cy + mag.r * Math.sin(a);
    s += `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" stroke="#8E8379" stroke-width="2.5" stroke-dasharray="7 6"/>`;
  });
  s += `<circle cx="${focus.cx}" cy="${focus.cy}" r="${focus.r}" fill="#FFFFFF" fill-opacity="0.16" stroke="#FFFFFF" stroke-width="4"/>`;
  s += `<circle cx="${mag.cx}" cy="${mag.cy + 8}" r="${mag.r + 4}" fill="#2E2622" opacity="0.22" filter="url(#shadow)"/>`;
  s += `<g clip-path="url(#lens)"><circle cx="${mag.cx}" cy="${mag.cy}" r="${mag.r}" fill="url(#air)"/>`;
  const E = (x, y, rx, ry, rot, fill, extra = '') => `<ellipse cx="${(mag.cx + x).toFixed(0)}" cy="${(mag.cy + y).toFixed(0)}" rx="${rx}" ry="${ry}" transform="rotate(${rot} ${(mag.cx + x).toFixed(0)} ${(mag.cy + y).toFixed(0)})" fill="${fill}"${extra}/>`;
  GRAINS.forEach(g => { s += E(g[0], g[1], g[2] + 11, g[3] + 11, g[4], '#46B7C6'); });   // плёнки воды вокруг зёрен
  BRIDGES.forEach(b => { s += E(b[0], b[1], b[2], b[3], b[4], '#46B7C6'); });           // мениски между зёрнами
  GRAINS.forEach(g => { s += E(g[0], g[1], g[2], g[3], g[4], g[5], ' stroke="#A9814F" stroke-width="2.5"'); });
  GRAINS.forEach(g => { s += E(g[0] - g[2] * 0.25, g[1] - g[3] * 0.3, g[2] * 0.35, g[3] * 0.22, g[4], '#FFFFFF', ' opacity="0.28"'); });
  s += `</g><circle cx="${mag.cx}" cy="${mag.cy}" r="${mag.r}" fill="none" stroke="#FFFFFF" stroke-width="10"/>`;
  s += `<circle cx="${mag.cx}" cy="${mag.cy}" r="${mag.r + 5}" fill="none" stroke="#CFC6BB" stroke-width="2"/>`;
  // опорные точки для подписей (маленькие точки-якоря)
  Object.values(SENSOR_ANCHORS).forEach(([x, y]) => { s += `<circle cx="${x}" cy="${y}" r="7" fill="#231C17" stroke="#FFFFFF" stroke-width="3"/>`; });
  s += `</svg>`;
  return s;
}

// 3. Частица глины с облаком ионов и солями в поровом растворе (кадр 900 × 640)
function claySVG() {
  const r = rng(7);
  const W = 900, H = 640, cx = 450, cy = 318;
  const plus = (x, y, rad, fill) => `<circle cx="${x.toFixed(0)}" cy="${y.toFixed(0)}" r="${rad}" fill="${fill}"/><rect x="${(x - rad * 0.55).toFixed(1)}" y="${(y - rad * 0.12).toFixed(1)}" width="${(rad * 1.1).toFixed(1)}" height="${(rad * 0.24).toFixed(1)}" rx="2" fill="#fff"/><rect x="${(x - rad * 0.12).toFixed(1)}" y="${(y - rad * 0.55).toFixed(1)}" width="${(rad * 0.24).toFixed(1)}" height="${(rad * 1.1).toFixed(1)}" rx="2" fill="#fff"/>`;
  const minus = (x, y, rad, fill) => `<circle cx="${x.toFixed(0)}" cy="${y.toFixed(0)}" r="${rad}" fill="${fill}"/><rect x="${(x - rad * 0.55).toFixed(1)}" y="${(y - rad * 0.12).toFixed(1)}" width="${(rad * 1.1).toFixed(1)}" height="${(rad * 0.24).toFixed(1)}" rx="2" fill="#fff"/>`;
  let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
  <defs>
    <radialGradient id="halo" cx="0.5" cy="0.5" r="0.5"><stop offset="0" stop-color="#008C9E" stop-opacity="0.30"/><stop offset="0.7" stop-color="#008C9E" stop-opacity="0.10"/><stop offset="1" stop-color="#008C9E" stop-opacity="0"/></radialGradient>
    <linearGradient id="top" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#B07045"/><stop offset="1" stop-color="#97593A"/></linearGradient>
  </defs>
  <ellipse cx="${cx}" cy="${cy}" rx="400" ry="290" fill="url(#halo)"/>`;
  // вода: светлые молекулы-точки
  for (let i = 0; i < 70; i++) {
    const a = r() * Math.PI * 2, d = 0.35 + r() * 0.65, x = cx + Math.cos(a) * 400 * d, y = cy + Math.sin(a) * 280 * d;
    s += `<circle cx="${x.toFixed(0)}" cy="${y.toFixed(0)}" r="${(3 + r() * 3).toFixed(1)}" fill="#7CC9D3" opacity="0.55"/>`;
  }
  // пластинка глины: боковая грань и верх
  const top = [[260, 300], [355, 238], [545, 238], [640, 300], [545, 362], [355, 362]];
  const side = [[260, 300], [355, 362], [545, 362], [640, 300], [640, 334], [545, 396], [355, 396], [260, 334]];
  const pts = a => a.map(p => p.join(',')).join(' ');
  s += `<polygon points="${pts(side)}" fill="#6E3A22"/><polygon points="${pts(top)}" fill="url(#top)"/>`;
  // отрицательный заряд на поверхности
  [[345, 280], [415, 270], [485, 270], [555, 280], [380, 318], [450, 312], [520, 318], [415, 345], [485, 345]].forEach(([x, y]) => {
    s += `<rect x="${x - 13}" y="${y - 3.5}" width="26" height="7" rx="3" fill="#F7EDE3" opacity="0.95"/>`;
  });
  // облако катионов
  for (let k = 0; k < 14; k++) {
    const a = k / 14 * Math.PI * 2 + 0.2, x = cx + Math.cos(a) * 285, y = cy + Math.sin(a) * 165;
    s += plus(x, y, 19, '#008C9E');
  }
  for (let k = 0; k < 8; k++) {
    const a = k / 8 * Math.PI * 2 + 0.6, x = cx + Math.cos(a) * 225, y = cy + Math.sin(a) * 118;
    s += plus(x, y, 14, '#2BA7B8');
  }
  // соли в поровом растворе (пары + и −) по углам
  [[95, 95], [800, 110], [90, 545], [810, 540], [150, 300], [760, 330]].forEach(([x, y], i) => {
    s += plus(x, y, 13, '#008C9E') + minus(x + 30 + (i % 2) * 6, y + 18, 13, '#C0562B');
  });
  s += `</svg>`;
  return s;
}

async function png(svg, scale = 2) {
  const m = svg.match(/width="(\d+)" height="(\d+)"/);
  const w = +m[1] * scale;
  const buf = await sharp(Buffer.from(svg), { density: 72 * scale }).resize({ width: w }).png().toBuffer();
  return buf;
}
const uri = buf => 'image/png;base64,' + buf.toString('base64');

async function illustrations() {
  const [cover, sensor, clay] = await Promise.all([png(coverSVG()), png(sensorSVG()), png(claySVG())]);
  return { cover: uri(cover), sensor: uri(sensor), clay: uri(clay), _buf: { cover, sensor, clay } };
}
module.exports = { illustrations, coverSVG, sensorSVG, claySVG, png, SENSOR, SENSOR_ANCHORS };

if (require.main === module) {
  const fs = require('fs');
  const out = process.argv[2] || '.';
  illustrations().then(I => { for (const k of Object.keys(I._buf)) fs.writeFileSync(`${out}/ill_${k}.png`, I._buf[k]); console.log('ok'); });
}
