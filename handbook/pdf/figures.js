"use strict";

const fs = require("fs");
const path = require("path");

const figDir = path.join(__dirname, "fig");
fs.mkdirSync(figDir, { recursive: true });

function fmt(n) {
  return String(Math.round(n * 10) / 10).replace(".", ",");
}

function pathOf(pts) {
  return pts.map((p, i) => `${i ? "L" : "M"}${p[0].toFixed(2)},${p[1].toFixed(2)}`).join(" ") + " Z";
}

function ellipsePts(w, h, n = 140) {
  const pts = [];
  for (let i = 0; i < n; i++) {
    const a = -Math.PI / 2 + (i / n) * Math.PI * 2;
    pts.push([(w / 2) * Math.cos(a), (h / 2) * Math.sin(a)]);
  }
  return pts;
}

function radialPts(contains, n = 180) {
  const pts = [];
  for (let i = 0; i < n; i++) {
    const ang = -Math.PI / 2 + (i / n) * Math.PI * 2;
    const co = Math.cos(ang);
    const si = Math.sin(ang);
    let lo = 0;
    let hi = 24;
    for (let k = 0; k < 26; k++) {
      const mid = (lo + hi) / 2;
      if (contains(co * mid, si * mid)) lo = mid;
      else hi = mid;
    }
    pts.push([co * lo, si * lo]);
  }
  return pts;
}

function bubblePts(upper, lower) {
  const d = upper.cy - lower.cy;
  const a = (d * d + upper.r * upper.r - lower.r * lower.r) / (2 * d);
  const x = Math.sqrt(Math.max(upper.r * upper.r - a * a, 0));
  const yk = upper.cy - a;
  const pts = [];
  const aL = Math.atan2(yk - upper.cy, -x);
  const aR = Math.atan2(yk - upper.cy, x);
  const upperEnd = aR - Math.PI * 2;
  for (let i = 0; i <= 90; i++) {
    const ang = aL + (upperEnd - aL) * (i / 90);
    pts.push([upper.r * Math.cos(ang), upper.cy + upper.r * Math.sin(ang)]);
  }
  const bL = Math.atan2(yk - lower.cy, -x);
  const bR = Math.atan2(yk - lower.cy, x);
  for (let i = 0; i <= 60; i++) {
    const ang = bR + (bL - bR) * (i / 60);
    pts.push([lower.r * Math.cos(ang), lower.cy + lower.r * Math.sin(ang)]);
  }
  return pts;
}

function bounds(pts) {
  let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
  for (const [x, y] of pts) {
    if (x < minX) minX = x;
    if (x > maxX) maxX = x;
    if (y < minY) minY = y;
    if (y > maxY) maxY = y;
  }
  return { minX, maxX, minY, maxY };
}

function insetPts(pts, wall) {
  return pts.map(([x, y]) => {
    const r = Math.hypot(x, y) || 1;
    const nr = Math.max(r - wall, 0.15);
    return [(x / r) * nr, (y / r) * nr];
  });
}

function layoutSeats(groups, span, aisle) {
  const n = groups.reduce((a, b) => a + b, 0);
  const nA = groups.length - 1;
  const seat = n ? (span - nA * aisle) / n : 0;
  let x = -(n * seat + nA * aisle) / 2;
  const seats = [];
  const groupsBox = [];
  const aisles = [];
  groups.forEach((g, gi) => {
    const x0 = x;
    for (let i = 0; i < g; i++) {
      seats.push({ x: x + seat / 2, w: seat });
      x += seat;
    }
    groupsBox.push({ x0, x1: x });
    if (gi < groups.length - 1) {
      aisles.push({ x: x + aisle / 2, w: aisle });
      x += aisle;
    }
  });
  return { seats, groupsBox, aisles, seat };
}

function sectionSVG(cfg) {
  const pts = cfg.pts;
  const b = bounds(pts);
  const inner = insetPts(pts, cfg.wall);
  const ppm = 78;
  const padL = 28;
  const padR = 54;
  const padT = 36;
  const padB = 42;
  const vbW = (b.maxX - b.minX) * ppm + padL + padR;
  const vbH = (b.maxY - b.minY) * ppm + padT + padB;
  const X = (x) => padL + (x - b.minX) * ppm;
  const Y = (y) => padT + (b.maxY - y) * ppm;
  const map = (list) => list.map(([x, y]) => [X(x), Y(y)]);
  const parts = [];
  parts.push(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${vbW.toFixed(1)} ${vbH.toFixed(1)}">`);
  parts.push(`<path fill="#7eafd4" stroke="#24384d" stroke-width="1.6" d="${pathOf(map(pts))}"/>`);
  parts.push(`<defs><clipPath id="in-${cfg.id}">${`<path d="${pathOf(map(inner))}"/>`}</clipPath></defs>`);
  parts.push(`<path fill="#f4f7f8" d="${pathOf(map(inner))}"/>`);
  parts.push(`<g clip-path="url(#in-${cfg.id})">`);

  cfg.decks.forEach((deck, di) => {
    const lay = layoutSeats(deck.groups, deck.span, deck.aisle);
    const yHigh = deck.floorY + deck.cabinH - 0.06;
    const drop = deck.cabinH < 2.05 ? 0.18 : 0.40;
    const wave = [];
    const steps = 64;
    for (let i = 0; i <= steps; i++) {
      const x = -deck.span / 2 + (deck.span * i) / steps;
      let y = yHigh;
      for (const g of lay.groupsBox) {
        if (x < g.x0 - 0.08 || x > g.x1 + 0.08) continue;
        const mid = (g.x0 + g.x1) / 2;
        const half = (g.x1 - g.x0) / 2 + 0.08;
        const u = Math.min(Math.abs(x - mid) / half, 1);
        y = yHigh - drop * Math.cos((u * Math.PI) / 2);
      }
      wave.push([x, y]);
    }
    const scallop = map(wave);
    const next = cfg.decks[di + 1];
    const crown = Y(next ? next.floorY - 0.06 : b.maxY);
    let d = `M${scallop[0][0].toFixed(2)},${scallop[0][1].toFixed(2)}`;
    scallop.slice(1).forEach((p) => { d += ` L${p[0].toFixed(2)},${p[1].toFixed(2)}`; });
    d += ` L${scallop[scallop.length - 1][0].toFixed(2)},${crown.toFixed(2)} L${scallop[0][0].toFixed(2)},${crown.toFixed(2)} Z`;
    parts.push(`<path fill="#7eafd4" d="${d}"/>`);

    const floorY = Y(deck.floorY);
    parts.push(`<line x1="${X(-deck.span / 2 - 0.2)}" y1="${floorY}" x2="${X(deck.span / 2 + 0.2)}" y2="${floorY}" stroke="#9aa3ad" stroke-width="2.4"/>`);

    lay.seats.forEach((s) => {
      const backW = s.w * 0.84;
      const backH = 0.78;
      const bot = deck.floorY + 0.02;
      const top = bot + backH;
      parts.push(`<rect x="${X(s.x - backW / 2)}" y="${Y(top)}" width="${backW * ppm}" height="${backH * ppm}" rx="4" fill="#8b9198"/>`);
      const hw = backW * 0.7;
      const hh = 0.16;
      parts.push(`<rect x="${X(s.x - hw / 2)}" y="${Y(top + hh - 0.02)}" width="${hw * ppm}" height="${hh * ppm}" rx="3" fill="#5e656c"/>`);
    });

    const aisle = lay.aisles.slice().sort((a, c) => Math.abs(a.x) - Math.abs(c.x))[0];
    if (aisle) person(parts, aisle.x, deck.floorY, X, Y, ppm);

    if (di === 0 && cfg.cargo === "bulk") {
      const holdW = Math.min(2.4, deck.span * 0.7);
      const holdH = 0.62;
      const holdTop = deck.floorY - 0.22;
      parts.push(`<rect x="${X(-holdW / 2)}" y="${Y(holdTop)}" width="${holdW * ppm}" height="${holdH * ppm}" rx="3" fill="#c5ccd3" stroke="#8d97a3" stroke-width="0.8"/>`);
      parts.push(`<text x="${X(0)}" y="${Y(holdTop - holdH * 0.62)}" text-anchor="middle" font-family="Liberation Sans, sans-serif" font-size="11" fill="#3e4852">навалом</text>`);
    }
    if (di === 0 && cfg.cargo === "boxes") {
      const boxH = Math.min(1.15, (deck.floorY - b.minY) * 0.55);
      const boxW = Math.min(1.55, deck.span / 4);
      const gap = 0.12;
      const by = deck.floorY - 0.08;
      [-1, 1].forEach((side) => {
        const cx = side * (boxW / 2 + gap / 2);
        parts.push(`<rect x="${X(cx - boxW / 2)}" y="${Y(by)}" width="${boxW * ppm}" height="${boxH * ppm}" rx="2" fill="#9aa1a8" stroke="#6d737a" stroke-width="0.8"/>`);
        parts.push(`<line x1="${X(cx - boxW / 2)}" y1="${Y(by - boxH * 0.72)}" x2="${X(cx + boxW / 2)}" y2="${Y(by - boxH * 0.72)}" stroke="#6d737a" stroke-width="0.7"/>`);
      });
    }
  });
  parts.push(`</g>`);

  const tagW = Math.max(36, cfg.tag.length * 7.2);
  const tagX = X(0) - tagW / 2;
  const tagY = 8;
  parts.push(`<rect x="${tagX}" y="${tagY}" width="${tagW}" height="16" rx="2" fill="#f7f8f6" stroke="#24384d" stroke-width="1"/>`);
  parts.push(`<text x="${X(0)}" y="${tagY + 12}" text-anchor="middle" font-family="Liberation Sans, sans-serif" font-size="10" fill="#24384d">${cfg.tag}</text>`);

  hDim(parts, X(b.minX), X(b.maxX), vbH - 16, `${fmt(b.maxX - b.minX)} м`);
  vDim(parts, vbW - 18, Y(b.maxY), Y(b.minY), `${fmt(b.maxY - b.minY)} м`);
  parts.push(`</svg>`);
  return parts.join("");
}

function person(parts, cx, floor, X, Y, ppm) {
  const xm = (dx) => X(cx + dx);
  const ym = (dy) => Y(floor + dy);
  parts.push(`<g>`);
  parts.push(`<rect x="${xm(-0.08)}" y="${ym(0.9)}" width="${0.06 * ppm}" height="${0.9 * ppm}" rx="2" fill="#243044"/>`);
  parts.push(`<rect x="${xm(0.02)}" y="${ym(0.9)}" width="${0.06 * ppm}" height="${0.9 * ppm}" rx="2" fill="#243044"/>`);
  parts.push(`<rect x="${xm(-0.13)}" y="${ym(1.48)}" width="${0.26 * ppm}" height="${0.58 * ppm}" rx="3" fill="#1f4b73"/>`);
  parts.push(`<circle cx="${xm(0)}" cy="${ym(1.62)}" r="${0.12 * ppm}" fill="#f0c7a4" stroke="#2a241f" stroke-width="0.6"/>`);
  parts.push(`</g>`);
}

function hDim(parts, x1, x2, y, text) {
  const a = Math.min(x1, x2);
  const b = Math.max(x1, x2);
  parts.push(`<line x1="${a}" y1="${y}" x2="${b}" y2="${y}" stroke="#1c2430" stroke-width="1"/>`);
  head(parts, a, y, 1, 0);
  head(parts, b, y, -1, 0);
  parts.push(`<text x="${(a + b) / 2}" y="${y - 4}" text-anchor="middle" font-family="Liberation Sans, sans-serif" font-size="11" fill="#1c2430">${text}</text>`);
}

function vDim(parts, x, y1, y2, text) {
  const a = Math.min(y1, y2);
  const b = Math.max(y1, y2);
  parts.push(`<line x1="${x}" y1="${a}" x2="${x}" y2="${b}" stroke="#1c2430" stroke-width="1"/>`);
  head(parts, x, a, 0, 1);
  head(parts, x, b, 0, -1);
  const ty = (a + b) / 2;
  parts.push(`<text transform="translate(${x + 12},${ty}) rotate(-90)" text-anchor="middle" font-family="Liberation Sans, sans-serif" font-size="11" fill="#1c2430">${text}</text>`);
}

function head(parts, x, y, dx, dy) {
  const s = 4.5;
  if (dx) {
    parts.push(`<polygon points="${x},${y} ${x + dx * s},${y - 2.6} ${x + dx * s},${y + 2.6}" fill="#1c2430"/>`);
  } else {
    parts.push(`<polygon points="${x},${y} ${x - 2.6},${y + dy * s} ${x + 2.6},${y + dy * s}" fill="#1c2430"/>`);
  }
}

function threeView(o) {
  const W = 760;
  const sideH = 168;
  const planH = 210;
  const gap = 28;
  const H = 18 + sideH + gap + planH + 16;
  const frontX = 560;
  const parts = [];
  parts.push(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}">`);
  parts.push(`<rect x="0" y="0" width="${W}" height="${H}" fill="#fbfcfd"/>`);
  parts.push(`<text x="8" y="14" font-family="Liberation Sans, sans-serif" font-size="12" fill="#16324f">${o.title} · метры</text>`);

  const sPpm = Math.min(500 / o.L, (sideH - 36) / o.H);
  sideView(parts, o, 16, 28, sPpm);

  const pPpm = Math.min(520 / o.L, (planH - 28) / o.span);
  planView(parts, o, 16, 28 + sideH + gap, pPpm);

  frontView(parts, o, frontX, 36, 188, sideH - 8);
  parts.push(`</svg>`);
  return parts.join("");
}

function sideView(parts, o, ox, oy, ppm) {
  const X = (x) => ox + x * ppm;
  const Y = (y) => oy + (o.H - y) * ppm;
  const gear = 2.2;
  const r = o.fuseH / 2;
  const belly = gear;
  const crown = gear + o.fuseH;
  const cy = gear + r;
  const x1 = o.L * 0.1;
  const x2 = o.L * 0.78;
  const humpEnd = o.L * 0.42;
  const humpTop = crown + r * 0.7;

  function crownAt(x) {
    if (!o.hump) return crown;
    if (x < humpEnd * 0.62) return humpTop;
    if (x < humpEnd) {
      const t = (x - humpEnd * 0.62) / (humpEnd * 0.38);
      return humpTop + (crown - humpTop) * t;
    }
    return crown;
  }

  const pts = [[0, cy]];
  pts.push([o.L * 0.035, cy + r * 0.45]);
  pts.push([x1, o.hump ? humpTop : crown]);
  for (let i = 1; i <= 28; i++) {
    const x = x1 + ((x2 - x1) * i) / 28;
    pts.push([x, crownAt(x)]);
  }
  pts.push([o.L * 0.9, cy + r * 0.25]);
  pts.push([o.L * 0.97, cy]);
  pts.push([o.L * 0.9, belly + r * 0.2]);
  pts.push([x2, belly]);
  pts.push([x1, belly]);
  pts.push([o.L * 0.04, cy - r * 0.2]);
  const d = pts.map((p, i) => `${i ? "L" : "M"}${X(p[0]).toFixed(1)},${Y(p[1]).toFixed(1)}`).join(" ") + " Z";
  parts.push(`<path d="${d}" fill="#f4f7fb" stroke="#1c2430" stroke-width="1.3"/>`);

  const winY = cy + r * 0.32;
  for (let x = x1 + 0.6; x < x2 - 1.2; x += o.L / 38) {
    parts.push(`<rect x="${X(x)}" y="${Y(winY + 0.16)}" width="${Math.max(1.4, 0.32 * ppm)}" height="${Math.max(1.6, 0.22 * ppm)}" rx="0.6" fill="#6f93ab"/>`);
  }
  parts.push(`<rect x="${X(o.L * 0.02)}" y="${Y(cy + r * 0.15)}" width="${0.7 * ppm}" height="${0.28 * ppm}" rx="1" fill="#6f93ab"/>`);

  const xWing = o.L * 0.33;
  const chord = o.L * 0.2;
  wingSide(parts, X, Y, xWing, chord, belly - r * 0.02, r * 0.42);

  if (o.engines === 4 || o.engines === 2) {
    const n = o.engines === 4 ? 2 : 1;
    for (let i = 0; i < n; i++) {
      const ex = xWing + chord * (0.22 + i * 0.5);
      nacelle(parts, X, Y, ex, belly - r * 0.72, chord * 0.34, r * 0.36);
    }
  }
  if (o.engines === 3) {
    parts.push(`<ellipse cx="${X(o.L * 0.8)}" cy="${Y(cy - r * 0.05)}" rx="${r * 0.85 * ppm}" ry="${r * 0.48 * ppm}" fill="#d5dde6" stroke="#1c2430" stroke-width="1.1"/>`);
  }

  const finBase = o.hump ? crown : crown;
  const fin = [
    [o.L * 0.7, finBase],
    [o.L * 0.84, finBase + r * 0.05],
    [o.L * 0.95, o.H - 0.2],
    [o.L * 0.78, o.H - 0.2],
  ];
  parts.push(`<polygon points="${fin.map((p) => `${X(p[0]).toFixed(1)},${Y(p[1]).toFixed(1)}`).join(" ")}" fill="#f4f7fb" stroke="#1c2430" stroke-width="1.2"/>`);
  if (o.engines === 3) {
    const ty = o.H - 0.55;
    parts.push(`<polygon points="${X(o.L * 0.7)},${Y(ty + 0.12)} ${X(o.L * 1.02)},${Y(ty + 0.12)} ${X(o.L * 1.02)},${Y(ty - 0.12)} ${X(o.L * 0.7)},${Y(ty - 0.12)}" fill="#f4f7fb" stroke="#1c2430" stroke-width="1.1"/>`);
    parts.push(`<ellipse cx="${X(o.L * 0.76)}" cy="${Y(crown + r * 0.35)}" rx="${r * 0.28 * ppm}" ry="${r * 0.22 * ppm}" fill="#d5dde6" stroke="#1c2430" stroke-width="0.9"/>`);
  } else {
    const stabY = cy + r * 0.05;
    parts.push(`<line x1="${X(o.L * 0.86)}" y1="${Y(stabY)}" x2="${X(o.L * 1.01)}" y2="${Y(stabY)}" stroke="#1c2430" stroke-width="2.4"/>`);
  }

  parts.push(`<line x1="${X(o.L * 0.12)}" y1="${Y(belly)}" x2="${X(o.L * 0.12)}" y2="${Y(0.15)}" stroke="#1c2430" stroke-width="1.2"/>`);
  parts.push(`<line x1="${X(o.L * 0.46)}" y1="${Y(belly)}" x2="${X(o.L * 0.46)}" y2="${Y(0.15)}" stroke="#1c2430" stroke-width="1.4"/>`);
  parts.push(`<line x1="${X(0)}" y1="${Y(0)}" x2="${X(o.L)}" y2="${Y(0)}" stroke="#c5ced6" stroke-width="1"/>`);
  dimH(parts, X(0), X(o.L), oy + o.H * ppm + 16, `длина ${fmt(o.L)}`);
  dimV(parts, X(o.L) + 16, Y(o.H), Y(0), `высота ${fmt(o.H)}`);
  parts.push(`<text x="${ox}" y="${oy + 12}" font-family="Liberation Sans, sans-serif" font-size="10" fill="#5c6b7a">сбоку</text>`);
}

function wingSide(parts, X, Y, x, chord, yMid, thick) {
  const pts = [
    [x, yMid],
    [x + chord * 0.15, yMid + thick * 0.45],
    [x + chord, yMid + thick * 0.05],
    [x + chord * 0.7, yMid - thick * 0.35],
  ];
  parts.push(`<polygon points="${pts.map((p) => `${X(p[0]).toFixed(1)},${Y(p[1]).toFixed(1)}`).join(" ")}" fill="#e7eef3" stroke="#1c2430" stroke-width="1.1"/>`);
}

function nacelle(parts, X, Y, x, y, len, rad) {
  parts.push(`<ellipse cx="${X(x)}" cy="${Y(y)}" rx="${(len / 2) * (X(1) - X(0))}" ry="${rad * (Y(0) - Y(1))}" fill="#d5dde6" stroke="#1c2430" stroke-width="1"/>`);
}

function planView(parts, o, ox, oy, ppm) {
  const X = (x) => ox + x * ppm;
  const Z = (z) => oy + 8 + (o.span / 2 - z) * ppm;
  parts.push(`<text x="${ox}" y="${oy + 4}" font-family="Liberation Sans, sans-serif" font-size="10" fill="#5c6b7a">сверху</text>`);
  const nose = o.L * 0.08;
  const tail = o.L * 0.12;
  const hw = o.fuseW / 2;
  const body = [
    [nose, 0],
    [o.L * 0.16, hw],
    [o.L - tail, hw],
    [o.L, 0],
    [o.L - tail, -hw],
    [o.L * 0.16, -hw],
  ];
  parts.push(`<polygon points="${body.map((p) => `${X(p[0]).toFixed(1)},${Z(p[1]).toFixed(1)}`).join(" ")}" fill="#f4f7fb" stroke="#1c2430" stroke-width="1.2"/>`);
  const xRoot = o.L * 0.34;
  const rootC = o.L * 0.22;
  const tipC = rootC * 0.22;
  const sweep = o.span * 0.22;
  const semi = o.span / 2;
  [-1, 1].forEach((s) => {
    const wing = [
      [xRoot, 0],
      [xRoot + sweep, s * semi],
      [xRoot + sweep + tipC, s * semi],
      [xRoot + rootC, 0],
    ];
    parts.push(`<polygon points="${wing.map((p) => `${X(p[0]).toFixed(1)},${Z(p[1]).toFixed(1)}`).join(" ")}" fill="#e7eef3" stroke="#1c2430" stroke-width="1.1"/>`);
  });
  if (o.engines === 3) {
    [-1, 1].forEach((s) => {
      parts.push(`<ellipse cx="${X(o.L * 0.84)}" cy="${Z(s * hw * 1.15)}" rx="${o.L * 0.035 * ppm}" ry="${hw * 0.45 * ppm}" fill="#d5dde6" stroke="#1c2430" stroke-width="0.8"/>`);
    });
  }
  if (o.engines === 2 || o.engines === 4) {
    const stations = o.engines === 4 ? [0.28, 0.46] : [0.34];
    stations.forEach((f) => {
      [-1, 1].forEach((s) => {
        const z = s * semi * f;
        const ex = xRoot + sweep * f + rootC * 0.25;
        parts.push(`<ellipse cx="${X(ex)}" cy="${Z(z)}" rx="${rootC * 0.18 * ppm}" ry="${hw * 0.55 * ppm}" fill="#d5dde6" stroke="#1c2430" stroke-width="0.9"/>`);
      });
    });
  }
  dimH(parts, X(0), X(o.L), oy + 8 + o.span * ppm + 14, `длина ${fmt(o.L)}`);
  dimV(parts, X(o.L) + 14, Z(o.span / 2), Z(-o.span / 2), `размах ${fmt(o.span)}`);
}

function frontView(parts, o, ox, oy, boxW, boxH) {
  parts.push(`<text x="${ox}" y="${oy - 4}" font-family="Liberation Sans, sans-serif" font-size="10" fill="#5c6b7a">спереди</text>`);
  const ppm = (boxH - 26) / o.H;
  const cx = ox + boxW / 2;
  const X = (x) => cx + x * ppm;
  const Y = (y) => oy + 4 + (o.H - y) * ppm;
  const gear = 2.2;
  const r = o.fuseH / 2;
  const cy = gear + r;
  const wingPpm = (boxW * 0.9) / o.span;
  const Xw = (x) => cx + x * wingPpm;
  const semi = o.span / 2;
  const wy = cy + r * 0.2;
  parts.push(`<polygon points="${Xw(-semi)},${Y(wy + 0.4)} ${Xw(0)},${Y(wy + r * 0.15)} ${Xw(semi)},${Y(wy + 0.4)} ${Xw(semi)},${Y(wy)} ${Xw(0)},${Y(wy + r * 0.05)} ${Xw(-semi)},${Y(wy)}" fill="#e7eef3" stroke="#1c2430" stroke-width="1"/>`);
  const engR = Math.max(r * 0.34, 0.7);
  if (o.engines === 2 || o.engines === 4) {
    const zs = o.engines === 4 ? [0.28, 0.5] : [0.34];
    zs.forEach((f) => {
      [-1, 1].forEach((s) => {
        parts.push(`<circle cx="${Xw(s * semi * f)}" cy="${Y(wy - engR * 0.2)}" r="${engR * ppm}" fill="#d5dde6" stroke="#1c2430" stroke-width="1"/>`);
      });
    });
  }
  parts.push(`<ellipse cx="${X(0)}" cy="${Y(cy)}" rx="${(o.fuseW / 2) * ppm}" ry="${r * ppm}" fill="#f4f7fb" stroke="#1c2430" stroke-width="1.3"/>`);
  if (o.hump) {
    parts.push(`<ellipse cx="${X(0)}" cy="${Y(cy + r * 0.85)}" rx="${(o.fuseW / 2) * 0.55 * ppm}" ry="${r * 0.55 * ppm}" fill="#f4f7fb" stroke="#1c2430" stroke-width="1.1"/>`);
  }
  const finTop = cy + r + Math.min(r * 1.15, o.H * 0.22);
  parts.push(`<polygon points="${X(0)},${Y(finTop)} ${X(-r * 0.16)},${Y(cy + r * 0.92)} ${X(r * 0.16)},${Y(cy + r * 0.92)}" fill="#f4f7fb" stroke="#1c2430" stroke-width="1"/>`);
  if (o.engines === 3) {
    [-1, 1].forEach((s) => {
      parts.push(`<circle cx="${X(s * (o.fuseW / 2 + engR * 0.15))}" cy="${Y(cy)}" r="${engR * 0.75 * ppm}" fill="#d5dde6" stroke="#1c2430" stroke-width="1"/>`);
    });
    parts.push(`<line x1="${Xw(-semi * 0.28)}" y1="${Y(o.H - 0.4)}" x2="${Xw(semi * 0.28)}" y2="${Y(o.H - 0.4)}" stroke="#1c2430" stroke-width="2.4"/>`);
  }
  parts.push(`<line x1="${X(-0.35)}" y1="${Y(gear)}" x2="${X(-0.35)}" y2="${Y(0.1)}" stroke="#1c2430" stroke-width="1.3"/>`);
  parts.push(`<line x1="${X(0.35)}" y1="${Y(gear)}" x2="${X(0.35)}" y2="${Y(0.1)}" stroke="#1c2430" stroke-width="1.3"/>`);
  parts.push(`<text x="${cx}" y="${oy + boxH - 2}" text-anchor="middle" font-family="Liberation Sans, sans-serif" font-size="10" fill="#1c2430">фюзеляж ${fmt(o.fuseW)} · размах ${fmt(o.span)}</text>`);
}

function dimH(parts, x1, x2, y, text) {
  parts.push(`<line x1="${x1}" y1="${y}" x2="${x2}" y2="${y}" stroke="#1c2430" stroke-width="0.9"/>`);
  parts.push(`<line x1="${x1}" y1="${y - 3}" x2="${x1}" y2="${y + 3}" stroke="#1c2430" stroke-width="0.9"/>`);
  parts.push(`<line x1="${x2}" y1="${y - 3}" x2="${x2}" y2="${y + 3}" stroke="#1c2430" stroke-width="0.9"/>`);
  parts.push(`<text x="${(x1 + x2) / 2}" y="${y - 3}" text-anchor="middle" font-family="Liberation Sans, sans-serif" font-size="10" fill="#1c2430">${text}</text>`);
}

function dimV(parts, x, y1, y2, text) {
  const a = Math.min(y1, y2);
  const b = Math.max(y1, y2);
  parts.push(`<line x1="${x}" y1="${a}" x2="${x}" y2="${b}" stroke="#1c2430" stroke-width="0.9"/>`);
  parts.push(`<text transform="translate(${x + 11},${(a + b) / 2}) rotate(-90)" text-anchor="middle" font-family="Liberation Sans, sans-serif" font-size="10" fill="#1c2430">${text}</text>`);
}

const narrow = bubblePts({ cy: 0.25, r: 1.88 }, { cy: -0.3, r: 1.633 });
const b747 = radialPts((x, y) => {
  const circles = [{ cy: -0.15, r: 3.25 }, { cy: 2.85, r: 1.7 }];
  return circles.some((c) => x * x + (y - c.cy) * (y - c.cy) <= c.r * c.r);
});

const sections = [
  {
    id: "narrow", tag: "707–757", wall: 0.18, pts: narrow, cargo: "bulk",
    decks: [{ floorY: -0.4, cabinH: 2.2, span: 3.4, aisle: 0.48, groups: [3, 3] }],
  },
  {
    id: "b767", tag: "767", wall: 0.2, pts: ellipsePts(5.03, 5.41), cargo: "boxes",
    decks: [{ floorY: -0.85, cabinH: 2.2, span: 4.55, aisle: 0.48, groups: [2, 3, 2] }],
  },
  {
    id: "b777", tag: "777", wall: 0.22, pts: ellipsePts(6.2, 6.2), cargo: "boxes",
    decks: [{ floorY: -0.85, cabinH: 2.4, span: 5.7, aisle: 0.48, groups: [3, 4, 3] }],
  },
  {
    id: "b787", tag: "787", wall: 0.2, pts: ellipsePts(5.77, 5.94), cargo: "boxes",
    decks: [{ floorY: -0.85, cabinH: 2.3, span: 5.3, aisle: 0.48, groups: [3, 3, 3] }],
  },
  {
    id: "b747", tag: "747", wall: 0.22, pts: b747, cargo: "boxes",
    decks: [
      { floorY: -0.9, cabinH: 2.4, span: 5.85, aisle: 0.48, groups: [3, 4, 3] },
      { floorY: 2.05, cabinH: 1.9, span: 3.15, aisle: 0.5, groups: [2, 2] },
    ],
  },
];

const views = [
  { file: "ga-707", title: "707-320B", L: 46.6, span: 44.4, H: 12.9, fuseW: 3.76, fuseH: 4.06, engines: 4 },
  { file: "ga-727", title: "727-200", L: 46.7, span: 32.9, H: 10.6, fuseW: 3.76, fuseH: 4.06, engines: 3 },
  { file: "ga-737", title: "737-800", L: 39.5, span: 35.8, H: 12.5, fuseW: 3.76, fuseH: 4.06, engines: 2 },
  { file: "ga-747", title: "747-400", L: 70.7, span: 64.4, H: 19.4, fuseW: 6.5, fuseH: 6.6, engines: 4, hump: true },
  { file: "ga-757", title: "757-200", L: 47.3, span: 38.0, H: 13.6, fuseW: 3.76, fuseH: 4.06, engines: 2 },
  { file: "ga-767", title: "767-300ER", L: 54.9, span: 47.6, H: 15.8, fuseW: 5.03, fuseH: 5.41, engines: 2 },
  { file: "ga-777", title: "777-300ER", L: 73.9, span: 64.8, H: 18.5, fuseW: 6.2, fuseH: 6.2, engines: 2 },
  { file: "ga-787", title: "787-9", L: 62.8, span: 60.1, H: 17.0, fuseW: 5.77, fuseH: 5.94, engines: 2 },
];

for (const s of sections) fs.writeFileSync(path.join(figDir, `section-${s.id}.svg`), sectionSVG(s));
for (const v of views) fs.writeFileSync(path.join(figDir, `${v.file}.svg`), threeView(v));
console.log("figures ok");
