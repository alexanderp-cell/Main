"use strict";

const hullCache = new Map();
let svgSeq = 0;

function escXml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[ch]));
}

function fmtM(n, digits) {
  const d = digits == null ? (Math.abs(n) >= 10 ? 1 : 2) : digits;
  return n.toFixed(d).replace(".", ",") + " м";
}

function bubbleJoin(items) {
  const upper = items[0];
  const lower = items[1];
  const d = upper.cy - lower.cy;
  const R = upper.r;
  const r = lower.r;
  const a = (d * d + R * R - r * r) / (2 * d);
  const x = Math.sqrt(Math.max(R * R - a * a, 0));
  return { upper, lower, x, yk: upper.cy - a };
}

function getHull(ac) {
  if (hullCache.has(ac.id)) return hullCache.get(ac.id);
  const h = ac.section.hull;
  let contains;
  let halfWidth;
  if (h.type === "bubble") {
    const g = bubbleJoin(h.items);
    contains = (x, y) => {
      const c = y >= g.yk ? g.upper : g.lower;
      const dx = x - (c.cx || 0);
      const dy = y - c.cy;
      return dx * dx + dy * dy <= c.r * c.r + 1e-6;
    };
    halfWidth = (y) => {
      const c = y >= g.yk ? g.upper : g.lower;
      const dy = y - c.cy;
      if (Math.abs(dy) > c.r) return 0;
      return Math.sqrt(Math.max(c.r * c.r - dy * dy, 0));
    };
  } else if (h.type === "ellipse") {
    const rx = h.w / 2;
    const ry = h.h / 2;
    contains = (x, y) => (x * x) / (rx * rx) + (y * y) / (ry * ry) <= 1.0005;
  } else {
    contains = (x, y) => h.items.some((c) => {
      const dx = x - (c.cx || 0);
      const dy = y - c.cy;
      return dx * dx + dy * dy <= c.r * c.r + 1e-5;
    });
  }
  function radiusAt(ang) {
    const coss = Math.cos(ang);
    const sinn = Math.sin(ang);
    let lo = 0;
    let hi = 30;
    if (!contains(0, 0)) return 0;
    for (let i = 0; i < 30; i++) {
      const mid = (lo + hi) / 2;
      if (contains(coss * mid, sinn * mid)) lo = mid;
      else hi = mid;
    }
    return lo;
  }
  const pts = [];
  let minX = Infinity;
  let maxX = -Infinity;
  let minY = Infinity;
  let maxY = -Infinity;
  const N = 240;
  for (let i = 0; i < N; i++) {
    const ang = -Math.PI / 2 + (i / N) * Math.PI * 2;
    const r = radiusAt(ang);
    const x = Math.cos(ang) * r;
    const y = Math.sin(ang) * r;
    pts.push([x, y]);
    if (x < minX) minX = x;
    if (x > maxX) maxX = x;
    if (y < minY) minY = y;
    if (y > maxY) maxY = y;
  }
  if (!halfWidth) {
    halfWidth = (y) => {
      if (y < minY || y > maxY) return 0;
      if (h.type === "ellipse") {
        const rx = h.w / 2;
        const ry = h.h / 2;
        const t = 1 - (y * y) / (ry * ry);
        return t <= 0 ? 0 : rx * Math.sqrt(t);
      }
      let m = 0;
      for (const c of h.items) {
        const dy = y - c.cy;
        if (Math.abs(dy) <= c.r) {
          const hw = Math.sqrt(Math.max(c.r * c.r - dy * dy, 0));
          const cx = c.cx || 0;
          m = Math.max(m, Math.abs(cx + hw), Math.abs(cx - hw));
        }
      }
      return m;
    };
  }
  const hull = { pts, minX, maxX, minY, maxY, contains, halfWidth };
  hullCache.set(ac.id, hull);
  return hull;
}

function insetPts(pts, wall) {
  return pts.map(([x, y]) => {
    const r = Math.hypot(x, y) || 1;
    const nr = Math.max(r - wall, 0.08);
    return [(x / r) * nr, (y / r) * nr];
  });
}

function pathOf(pts, X, Y) {
  return pts
    .map((p, i) => `${i ? "L" : "M"}${X(p[0]).toFixed(2)},${Y(p[1]).toFixed(2)}`)
    .join(" ") + " Z";
}

function seatLayout(deck) {
  const groups = deck.groups;
  const n = groups.reduce((a, b) => a + b, 0);
  const nA = groups.length - 1;
  const span = deck.span;
  let aisle = deck.aisle;
  let seat = n ? (span - nA * aisle) / n : 0;
  if (seat < 0.4 && nA) {
    seat = 0.4;
    aisle = (span - n * seat) / nA;
  }
  const total = n * seat + nA * aisle;
  let x = -total / 2;
  const seats = [];
  const aisles = [];
  const groupBoxes = [];
  groups.forEach((g, gi) => {
    const start = x;
    for (let i = 0; i < g; i++) {
      seats.push({ x: x + seat / 2, w: seat, group: gi });
      x += seat;
    }
    groupBoxes.push({ x1: start, x2: x, group: gi });
    if (gi < groups.length - 1) {
      aisles.push({ x: x + aisle / 2, w: aisle, index: gi });
      x += aisle;
    }
  });
  return { seats, aisles, seat, aisle, total, groupBoxes };
}

function centralAisle(layout) {
  if (!layout.aisles.length) return null;
  return layout.aisles.slice().sort((a, b) => Math.abs(a.x) - Math.abs(b.x))[0];
}

function crossSectionSVG(ac, opts) {
  const mode = (opts && opts.mode) || "detail";
  const hull = getHull(ac);
  const wall = ac.section.wall == null ? 0.13 : ac.section.wall;
  const spanM = hull.maxX - hull.minX;
  const heightM = hull.maxY - hull.minY;
  let ppm = opts && opts.ppm;
  if (!ppm) {
    if (mode === "detail") ppm = 520 / Math.max(spanM, heightM);
    else if (mode === "gallery") ppm = 48;
    else if (mode === "compare") ppm = 58;
    else ppm = 16;
  }
  const pad = mode === "detail" ? 86 : mode === "mini" ? 6 : 14;
  const seq = ++svgSeq;
  const vbW = spanM * ppm + pad * 2;
  const vbH = heightM * ppm + pad * 2;
  const X = (x) => pad + (x - hull.minX) * ppm;
  const Y = (y) => pad + (hull.maxY - y) * ppm;
  const inner = insetPts(hull.pts, wall);
  const clip = `cs-${ac.id}-${seq}`;
  const parts = [];
  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${vbW.toFixed(1)} ${vbH.toFixed(1)}" width="${vbW.toFixed(0)}" height="${vbH.toFixed(0)}" role="img" aria-label="${escXml(ac.name)}: поперечное сечение">`
  );
  parts.push(`<defs><clipPath id="${clip}">${pathOf(inner, X, Y)}</clipPath></defs>`);
  parts.push(`<path class="skin" d="${pathOf(hull.pts, X, Y)}"/>`);
  parts.push(`<path class="cavity" d="${pathOf(inner, X, Y)}"/>`);

  const decks = ac.section.decks;
  const lowest = decks.reduce((m, d) => Math.min(m, d.floorY), Infinity);
  parts.push(
    `<rect class="belly" clip-path="url(#${clip})" x="${X(hull.minX)}" y="${Y(lowest)}" width="${spanM * ppm}" height="${(lowest - hull.minY) * ppm + 4}"/>`
  );
  if (mode === "detail" && lowest - hull.minY > 0.85) {
    parts.push(
      `<text class="deck-label" text-anchor="middle" x="${X(0)}" y="${Y((lowest + hull.minY) / 2)}">багажный отсек</text>`
    );
  }

  decks.forEach((deck) => {
    const hw = hull.halfWidth(deck.floorY) - wall * 0.4;
    const floorHalf = Math.min(hw, Math.max(deck.span / 2 + 0.08, deck.cabinW ? deck.cabinW / 2 : deck.span / 2));
    const x1 = -floorHalf;
    const x2 = floorHalf;
    const slab = 0.1;
    parts.push(
      `<rect class="floor" x="${X(x1)}" y="${Y(deck.floorY)}" width="${(x2 - x1) * ppm}" height="${slab * ppm}"/>`
    );
    const layout = seatLayout(deck);
    const aisle = centralAisle(layout);
    drawBins(parts, deck, layout, hull, wall, X, Y, ppm, mode);
    layout.seats.forEach((s) => {
      parts.push(seatSVG(s, deck.floorY, X, Y, ppm));
    });
    if (aisle) {
      parts.push(personSVG(aisle.x, deck.floorY, 1.75, X, Y, ppm));
      if (mode === "detail") {
        const ceil = deck.floorY + deck.cabinH;
        parts.push(
          `<line class="ceil" x1="${X(aisle.x - aisle.w / 2)}" y1="${Y(ceil)}" x2="${X(aisle.x + aisle.w / 2)}" y2="${Y(ceil)}"/>`
        );
      }
    }
    drawWindows(parts, deck, hull, X, Y, ppm);
    if (mode === "detail" && decks.length > 1) {
      parts.push(
        `<text class="deck-label" x="${X(hull.minX + 0.18)}" y="${Y(deck.floorY + deck.cabinH * 0.55)}">${escXml(deck.label)}</text>`
      );
    }
  });

  if (mode === "detail") {
    drawDimensions(parts, ac, hull, X, Y, ppm, pad, vbW, vbH);
  }

  parts.push("</svg>");
  return parts.join("");
}

function drawWindows(parts, deck, hull, X, Y, ppm) {
  const y = deck.floorY + Math.min(1.12, deck.cabinH * 0.55);
  const hw = hull.halfWidth(y);
  if (hw < 0.4) return;
  const wh = Math.min(0.36, deck.cabinH * 0.18) * ppm;
  const ww = Math.max(0.07, 0.09) * ppm;
  [-1, 1].forEach((side) => {
    const cx = side * (hw - 0.03);
    parts.push(
      `<rect class="window" x="${X(cx) - (side > 0 ? ww : 0)}" y="${Y(y) - wh / 2}" width="${ww}" height="${wh}" rx="1.5"/>`
    );
  });
}

function drawBins(parts, deck, layout, hull, wall, X, Y, ppm, mode) {
  if (mode === "mini") return;
  if (!layout.groupBoxes.length) return;
  const ceil = deck.floorY + deck.cabinH;
  [layout.groupBoxes[0], layout.groupBoxes[layout.groupBoxes.length - 1]].forEach((box) => {
    if (!box) return;
    const outerX = box.x1 < 0 ? box.x1 : box.x2;
    const innerX = box.x1 < 0 ? box.x2 : box.x1;
    const botOuter = deck.floorY + Math.min(1.5, deck.cabinH - 0.34);
    const botInner = ceil - 0.07;
    if (botInner - botOuter < 0.12) return;
    const top = ceil - 0.03;
    const pts = [
      [outerX, botOuter],
      [outerX, Math.min(top, yLimit(hull, outerX, wall, top))],
      [innerX, Math.min(top, yLimit(hull, innerX, wall, top))],
      [innerX, botInner],
    ];
    parts.push(`<path class="bin" d="${pathOf(pts, X, Y)}"/>`);
  });
}

function yLimit(hull, x, wall, want) {
  const hw = hull.halfWidth(want) - wall;
  if (Math.abs(x) <= hw) return want;
  return want - 0.12;
}

function seatSVG(s, floor, X, Y, ppm) {
  const w = s.w * 0.86;
  const backW = s.w * 0.62;
  const cushionY = floor + 0.43;
  const cushionH = 0.09;
  const backTop = floor + 1.16;
  const backBot = floor + 0.46;
  const armH = 0.16;
  const armW = Math.min(0.05, s.w * 0.12);
  const cx = s.x;
  return [
    `<g class="seat">`,
    `<rect class="arm" x="${X(cx - w / 2)}" y="${Y(cushionY + armH)}" width="${armW * ppm}" height="${armH * ppm}" rx="1"/>`,
    `<rect class="arm" x="${X(cx + w / 2) - armW * ppm}" y="${Y(cushionY + armH)}" width="${armW * ppm}" height="${armH * ppm}" rx="1"/>`,
    `<rect class="back" x="${X(cx - backW / 2)}" y="${Y(backTop)}" width="${backW * ppm}" height="${(backTop - backBot) * ppm}" rx="2"/>`,
    `<rect class="cushion" x="${X(cx - w / 2)}" y="${Y(cushionY)}" width="${w * ppm}" height="${cushionH * ppm}" rx="1.5"/>`,
    `</g>`,
  ].join("");
}

function personSVG(cx, floor, h, X, Y, ppm) {
  const xm = (dx) => X(cx + dx);
  const ym = (dy) => Y(floor + dy);
  const headR = h * 0.07;
  const headC = h - headR * 0.92;
  const neck = h * 0.835;
  const hip = h * 0.5;
  const sh = 0.2;
  const waist = 0.115;
  const torso = [
    [xm(-waist), ym(hip)],
    [xm(-sh), ym(neck - 0.04)],
    [xm(-0.07), ym(neck + 0.03)],
    [xm(0.07), ym(neck + 0.03)],
    [xm(sh), ym(neck - 0.04)],
    [xm(waist), ym(hip)],
  ];
  const d = torso.map((p, i) => `${i ? "L" : "M"}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join(" ") + " Z";
  const leg = (x0, x1) =>
    `<rect class="leg" x="${xm(x0)}" y="${ym(hip)}" width="${(xm(x1) - xm(x0)).toFixed(1)}" height="${(ym(0.02) - ym(hip)).toFixed(1)}" rx="2"/>`;
  return [
    `<g class="person">`,
    leg(-0.115, -0.02),
    leg(0.02, 0.115),
    `<path class="torso" d="${d}"/>`,
    `<circle class="head" cx="${xm(0)}" cy="${ym(headC)}" r="${(headR * ppm).toFixed(1)}"/>`,
    `<path class="hair" d="M ${xm(-headR * 0.86)} ${ym(headC)} A ${(headR * ppm).toFixed(1)} ${(headR * ppm).toFixed(1)} 0 0 1 ${xm(headR * 0.86)} ${ym(headC)} L ${xm(headR * 0.55)} ${ym(headC + headR * 0.15)} A ${(headR * 0.7 * ppm).toFixed(1)} ${(headR * 0.45 * ppm).toFixed(1)} 0 0 0 ${xm(-headR * 0.55)} ${ym(headC + headR * 0.15)} Z"/>`,
    `</g>`,
  ].join("");
}

function drawDimensions(parts, ac, hull, X, Y, ppm, pad, vbW, vbH) {
  const s = ac.section;
  const yBottom = Y(hull.minY) + 36;
  hDim(parts, X(hull.minX), X(hull.maxX), yBottom, labelOf(s.dimW, s.dimWKind, "фюзеляж"), vbW);

  const main = s.decks[0];
  const cabinY = Y(hull.maxY) - 28;
  const half = (main.cabinW || main.span) / 2;
  hDim(parts, X(-half), X(half), cabinY, labelOf(main.cabinW || main.span, main.cabinWKind || "published", "салон"), vbW);

  s.decks.forEach((deck, i) => {
    const x = X(hull.maxX) + 28 + (s.decks.length > 1 && i > 0 ? 8 : 0);
    const y1 = Y(deck.floorY);
    const y2 = Y(deck.floorY + deck.cabinH);
    const kind = deck.cabinHKind || "schematic";
    const text = fmtM(deck.cabinH) + (kind === "published" ? "" : " схема");
    vDim(parts, x, y1, y2, text);
  });

  const aisle = centralAisle(seatLayout(main));
  if (aisle) {
    const x = X(aisle.x) - 22;
    vDim(parts, x, Y(main.floorY), Y(main.floorY + 1.75), "175 см");
  }

  if (s.decks.length > 1) {
    const upper = s.decks[1];
    const uh = (upper.cabinW || upper.span) / 2;
    const uy = Y(upper.floorY + upper.cabinH) - 16;
    if (uy > 20) {
      hDim(parts, X(-uh), X(uh), uy, labelOf(upper.cabinW || upper.span, upper.cabinWKind || "schematic", upper.label), vbW);
    }
  }
}

function labelOf(meters, kind, prefix) {
  const base = `${prefix} ${fmtM(meters)}`;
  if (kind === "schematic") return base + " · схема";
  if (kind === "assumed") return base + " · допущение";
  return base;
}

function hDim(parts, x1, x2, y, text, vbW) {
  const a = Math.min(x1, x2);
  const b = Math.max(x1, x2);
  parts.push(`<g class="dim">`);
  parts.push(`<line x1="${a}" y1="${y}" x2="${b}" y2="${y}"/>`);
  parts.push(`<line x1="${a}" y1="${y - 5}" x2="${a}" y2="${y + 5}"/>`);
  parts.push(`<line x1="${b}" y1="${y - 5}" x2="${b}" y2="${y + 5}"/>`);
  const tx = Math.max(8, Math.min((a + b) / 2, (vbW || 2000) - 8));
  parts.push(`<text x="${tx}" y="${y - 7}" text-anchor="middle">${escXml(text)}</text>`);
  parts.push(`</g>`);
}

function vDim(parts, x, y1, y2, text) {
  const a = Math.min(y1, y2);
  const b = Math.max(y1, y2);
  parts.push(`<g class="dim">`);
  parts.push(`<line x1="${x}" y1="${a}" x2="${x}" y2="${b}"/>`);
  parts.push(`<line x1="${x - 5}" y1="${a}" x2="${x + 5}" y2="${a}"/>`);
  parts.push(`<line x1="${x - 5}" y1="${b}" x2="${x + 5}" y2="${b}"/>`);
  const ty = (a + b) / 2;
  parts.push(
    `<text transform="translate(${x + 8},${ty}) rotate(-90)" text-anchor="middle">${escXml(text)}</text>`
  );
  parts.push(`</g>`);
}

function gearSVG(ac) {
  const g = ac.gear;
  const r = 12;
  const gapX = 8;
  const gapY = 8;
  function bogieSize(n) {
    const cols = 2;
    const rows = Math.max(1, Math.round(n / 2));
    return { cols, rows, w: cols * r * 2 + (cols - 1) * gapX, h: rows * r * 2 + (rows - 1) * gapY };
  }
  const sizes = g.bogies.map(bogieSize);
  const gapB = 36;
  const gearW = sizes.reduce((s, b) => s + b.w, 0) + gapB * (sizes.length - 1);
  const maxH = Math.max.apply(null, sizes.map((b) => b.h));
  const noseR = 8;
  const noseW = noseR * 4 + 6;
  const pad = 36;
  const top = 18;
  const noseH = noseR * 2;
  const midGap = 36;
  const vbW = Math.max(gearW, noseW) + pad * 2 + 24;
  const vbH = top + noseH + midGap + maxH + 62;
  const cx = vbW / 2;
  const parts = [];
  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${vbW} ${vbH}" width="${vbW}" height="${vbH}" role="img" aria-label="Схема шасси ${escXml(ac.name)}">`
  );
  const bodyW = gearW + 70;
  const bodyH = vbH - 24;
  parts.push(
    `<rect class="gear-body" x="${cx - bodyW / 2}" y="8" width="${bodyW}" height="${bodyH}" rx="${Math.min(40, bodyW / 2)}"/>`
  );
  parts.push(`<line class="gear-cl" x1="${cx}" y1="16" x2="${cx}" y2="${vbH - 28}"/>`);
  const noseY = top + noseR;
  drawAxle(parts, cx, noseY, noseR, "nose");
  parts.push(`<text class="gear-label" x="${cx}" y="${noseY + noseR + 14}" text-anchor="middle">носовая · ${g.nose}</text>`);
  let x = cx - gearW / 2;
  const baseY = top + noseH + midGap;
  g.bogies.forEach((n, i) => {
    const sz = sizes[i];
    const ox = x;
    for (let row = 0; row < sz.rows; row++) {
      for (let col = 0; col < sz.cols; col++) {
        const wx = ox + r + col * (r * 2 + gapX);
        const wy = baseY + r + row * (r * 2 + gapY);
        parts.push(`<circle class="tire" cx="${wx}" cy="${wy}" r="${r}"/>`);
        parts.push(`<circle class="hub" cx="${wx}" cy="${wy}" r="${r * 0.38}"/>`);
      }
    }
    const label = (g.labels && g.labels[i]) || "опора";
    parts.push(
      `<text class="gear-label" x="${ox + sz.w / 2}" y="${baseY + maxH + 16}" text-anchor="middle">${escXml(label)}</text>`
    );
    parts.push(
      `<text class="gear-label" x="${ox + sz.w / 2}" y="${baseY + maxH + 30}" text-anchor="middle">${n}</text>`
    );
    x += sz.w + gapB;
  });
  parts.push(`<text class="gear-caption" x="${cx}" y="${vbH - 4}" text-anchor="middle">след шасси, вид сверху · не чертёж колеи</text>`);
  parts.push("</svg>");
  return parts.join("");
}

function drawAxle(parts, cx, cy, r, kind) {
  [-1, 1].forEach((side) => {
    const wx = cx + side * (r + 4);
    parts.push(`<circle class="tire ${kind || ""}" cx="${wx}" cy="${cy}" r="${r}"/>`);
    parts.push(`<circle class="hub" cx="${wx}" cy="${cy}" r="${r * 0.38}"/>`);
  });
}

function primerGearSVG() {
  const w = 360;
  const h = 220;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" role="img" aria-label="Принципиальная схема стойки с тележкой">
    <text class="primer-title" x="16" y="22">Принципиальная стойка с тележкой</text>
    <line class="oleo" x1="168" y1="36" x2="168" y2="108"/>
    <rect class="bogie-beam" x="78" y="104" width="180" height="12" rx="3"/>
    <g class="person">
      <circle class="tire" cx="110" cy="156" r="28"/>
      <circle class="hub" cx="110" cy="156" r="10"/>
      <circle class="tire" cx="226" cy="156" r="28"/>
      <circle class="hub" cx="226" cy="156" r="10"/>
    </g>
    <line class="axle" x1="110" y1="116" x2="110" y2="156"/>
    <line class="axle" x1="226" y1="116" x2="226" y2="156"/>
    <text class="gear-label" x="188" y="78">амортизатор</text>
    <text class="gear-label" x="188" y="100">тележка</text>
    <text class="gear-label" x="70" y="206">колесо</text>
    <text class="gear-label" x="196" y="206">колесо</text>
    <text class="gear-caption" x="16" y="214">Вид сбоку, одна ось за другой. Тормоз — на основных колёсах, не на рисунке.</text>
  </svg>`;
}

function abreastLabel(ac) {
  return ac.section.decks
    .map((d) => (d.label && ac.section.decks.length > 1 ? `${d.label} ${d.groups.join("+")}` : d.groups.join("+")))
    .join(" · ");
}

function validateHandbook() {
  const issues = [];
  if (typeof AIRCRAFT === "undefined") return ["AIRCRAFT missing"];
  const ids = new Set();
  AIRCRAFT.forEach((ac) => {
    if (ids.has(ac.id)) issues.push("dup " + ac.id);
    ids.add(ac.id);
    const hull = getHull(ac);
    if (!hull.contains(0, 0)) issues.push(ac.id + " origin outside hull");
    const width = hull.maxX - hull.minX;
    if (ac.section.dimW && Math.abs(width - ac.section.dimW) > 0.12) {
      issues.push(`${ac.id} width drawn ${width.toFixed(2)} vs dim ${ac.section.dimW}`);
    }
    ac.section.decks.forEach((d) => {
      if (d.cabinH < 1.75) issues.push(ac.id + " cabin shorter than person " + d.label);
      const top = d.floorY + d.cabinH;
      const wall = ac.section.wall == null ? 0.13 : ac.section.wall;
      if (top > hull.maxY - wall - 0.02) issues.push(ac.id + " ceiling clashes with skin " + d.label + " top " + top.toFixed(2) + " limit " + (hull.maxY - wall - 0.02).toFixed(2));
      if (d.floorY < hull.minY + 0.15) issues.push(ac.id + " floor too low " + d.label);
      const hw = hull.halfWidth(d.floorY + 0.05);
      if (d.span / 2 > hw - 0.02) {
        issues.push(`${ac.id} span ${d.span} exceeds floor width ${(hw * 2).toFixed(2)} ${d.label || ""}`);
      }
      if (!hull.contains(0, d.floorY + 0.2)) issues.push(ac.id + " floor not inside");
    });
    ["detail", "gallery", "mini"].forEach((mode) => {
      const svg = crossSectionSVG(ac, { mode });
      if (svg.includes("NaN")) issues.push(ac.id + " NaN " + mode);
    });
    const g = gearSVG(ac);
    if (g.includes("NaN")) issues.push(ac.id + " gear NaN");
  });
  return issues;
}
