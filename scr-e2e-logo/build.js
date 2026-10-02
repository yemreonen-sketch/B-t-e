// Builds the SCR E2E logo proposal as an editable PPTX (native shapes only)
// and, from the same geometry, SVG previews that are rendered to PNG for the upload form.
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5 in
pres.title = 'SCR E2E Logo Proposal';
pres.theme = { headFontFace: 'Arial', bodyFontFace: 'Arial' };

const RED = 'E1000F', DEEP = '8C0A1A', CORAL = 'FF5A4E', INK = '2B2B2B', GREY = '6B6B6B', LIGHT = 'F4F4F4', WHITE = 'FFFFFF';
const PALETTE = [RED, INK, CORAL, DEEP]; // four people, one ring
const rad = d => d * Math.PI / 180;
const K = 96; // SVG units: 1 in = 96 px, so HTML text inside foreignObject needs no scaling
const n4 = n => +(n * K).toFixed(2);
const norm = a => ((a % 360) + 360) % 360;

// ---------- dual backend ----------
const svgs = [];
function newSlide(bg) {
  const sl = pres.addSlide(); sl.background = { color: bg };
  const svg = [`<rect width="${n4(13.333)}" height="${n4(7.5)}" fill="#${bg}"/>`]; svgs.push(svg);
  return {
    rect(x, y, w, h, c, name) {
      sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: c }, line: { color: c, width: 0 }, objectName: name });
      svg.push(`<rect x="${n4(x)}" y="${n4(y)}" width="${n4(w)}" height="${n4(h)}" fill="#${c}"/>`);
    },
    oval(cx, cy, r, c, name) {
      sl.addShape(pres.shapes.OVAL, { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, fill: { color: c }, line: { color: c, width: 0 }, objectName: name });
      svg.push(`<circle cx="${n4(cx)}" cy="${n4(cy)}" r="${n4(r)}" fill="#${c}"/>`);
    },
    arc(cx, cy, R, t, a1, a2, c, name) { // clockwise from a1 to a2 (0 deg = 3 o'clock)
      sl.addShape(pres.shapes.BLOCK_ARC, { x: cx - R, y: cy - R, w: 2 * R, h: 2 * R, angleRange: [norm(a1), norm(a2)],
        arcThicknessRatio: t / R, fill: { color: c }, line: { color: c, width: 0 }, objectName: name });
      const r2 = R - t, p = (r, a) => `${n4(cx + r * Math.cos(rad(a)))} ${n4(cy + r * Math.sin(rad(a)))}`;
      const lg = norm(a2 - a1) > 180 ? 1 : 0;
      svg.push(`<path d="M${p(R, a1)} A${n4(R)} ${n4(R)} 0 ${lg} 1 ${p(R, a2)} L${p(r2, a2)} A${n4(r2)} ${n4(r2)} 0 ${lg} 0 ${p(r2, a1)} Z" fill="#${c}"/>`);
    },
    arrow(cx, cy, base, len, rot, c, name) { // isosceles triangle, apex up before rotation
      sl.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: cx - base / 2, y: cy - len / 2, w: base, h: len, rotate: rot,
        fill: { color: c }, line: { color: c, width: 0 }, objectName: name });
      svg.push(`<polygon points="${n4(cx)},${n4(cy - len / 2)} ${n4(cx + base / 2)},${n4(cy + len / 2)} ${n4(cx - base / 2)},${n4(cy + len / 2)}" transform="rotate(${rot} ${n4(cx)} ${n4(cy)})" fill="#${c}"/>`);
    },
    text(runs, o) {
      if (typeof runs === 'string') runs = [{ text: runs }];
      const opts = Object.assign({ margin: 0, isTextBox: true, fontFace: 'Arial' }, o);
      sl.addText(runs.map(r => ({ text: r.text, options: Object.assign({}, r.options || {}) })), opts);
      // SVG: HTML inside foreignObject (handles wrapping); 1in = 96 css px, so scale inner box
      const va = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[o.valign || 'top'];
      const paras = []; let p = [];
      runs.forEach(r => { p.push(r); if (r.options && r.options.breakLine) { paras.push(p); p = []; } }); if (p.length) paras.push(p);
      const html = paras.map(pp => `<div style="margin-bottom:${(o.paraSpaceAfter || 0) * 96 / 72}px">` + pp.map(r => {
        const ro = r.options || {};
        return `<span style="color:#${ro.color || o.color};font-weight:${(ro.bold ?? o.bold) ? 700 : 400}">${r.text.replace(/&/g, '&amp;')}</span>`;
      }).join('') + `</div>`).join('');
      svg.push(`<foreignObject x="${n4(o.x)}" y="${n4(o.y)}" width="${n4(o.w)}" height="${n4(o.h)}"><div xmlns="http://www.w3.org/1999/xhtml" style="width:${n4(o.w)}px;height:${n4(o.h)}px;display:flex;flex-direction:column;justify-content:${va};text-align:${o.align || 'left'};font-family:Arial,'Liberation Sans',sans-serif;font-size:${o.fontSize * 96 / 72}px;line-height:1.2;letter-spacing:${(o.charSpacing || 0) * 96 / 72}px;font-style:${o.italic ? 'italic' : 'normal'};font-weight:${o.bold ? 700 : 400};color:#${o.color || INK}">${html}</div></foreignObject>`);
    },
  };
}

// ---------- the mark ----------
// Four people link arms into one ring; the arrowhead makes it a moving journey; E2E sits at the heart.
function drawMark(s, cx, cy, R, o = {}) {
  const cols = o.colors || PALETTE, txt = o.text || INK, acc = o.accent || RED;
  const t = 0.27 * R;               // ring (arms) thickness
  const rh = 0.19 * R;              // head radius
  const dh = R + 0.05 * R + rh;     // head distance from centre
  const people = [225, 315, 45, 135];
  const half = 42.5;                  // arms span +-42.5 deg -> 5 deg "handshake" gaps
  people.forEach((a, i) => {
    const end = a === 225 ? a + half - 11 : a + half; // room for the arrowhead at 12 o'clock
    s.arc(cx, cy, R, t, a - half, end, cols[i], `Person ${i + 1} - arms`);
    s.oval(cx + dh * Math.cos(rad(a)), cy + dh * Math.sin(rad(a)), rh, cols[i], `Person ${i + 1} - head`);
  });
  const rm = R - t / 2, ang = 264;
  s.arrow(cx + rm * Math.cos(rad(ang)), cy + rm * Math.sin(rad(ang)), 1.85 * t, 1.1 * t, 90, cols[0], 'Journey arrow');
  if (o.center !== false) {
    s.text([{ text: 'E', options: { color: txt } }, { text: '2', options: { color: acc } }, { text: 'E', options: { color: txt } }],
      { x: cx - 0.7 * R, y: cy - 0.35 * R, w: 1.4 * R, h: 0.7 * R, align: 'center', valign: 'middle', bold: true,
        fontSize: Math.round(R * 72 * 0.40), color: txt, objectName: 'E2E' });
  }
}

function wordmark(s, x, y, w, size, o = {}) {
  const c = o.color || INK, acc = o.accent || RED, align = o.align || 'center', h = size / 72 * 1.25;
  s.text([{ text: 'SCR ', options: { color: c } }, { text: 'E', options: { color: c } }, { text: '2', options: { color: acc } }, { text: 'E', options: { color: c } }],
    { x, y, w, h, align, valign: 'middle', fontSize: size, bold: true, charSpacing: 2, color: c, objectName: 'Wordmark' });
  s.text('ONE JOURNEY. ONE IDENTITY.', { x, y: y + h + 0.04, w, h: size / 72 * 0.6, align, valign: 'top',
    fontSize: o.tagSize || Math.round(size * 0.36), color: o.tagColor || GREY, charSpacing: o.tagSpacing ?? 4, objectName: 'Tagline' });
}

// ---------- Slide 1: the logo ----------
let s = newSlide(WHITE);
drawMark(s, 13.333 / 2, 2.85, 1.75);
wordmark(s, 2.67, 5.35, 8, 40);

// ---------- Slide 2: submission board (logo + meaning) ----------
s = newSlide(WHITE);
s.rect(0, 0, 5.6, 7.5, LIGHT, 'Logo panel');
drawMark(s, 2.8, 3.0, 1.45);
wordmark(s, 0.4, 5.15, 4.8, 30);
s.text('SCR E2E Logo Proposal', { x: 6.2, y: 0.6, w: 6.6, h: 0.7, fontSize: 32, bold: true, color: INK });
s.text('People & Journey Together', { x: 6.2, y: 1.3, w: 6.6, h: 0.45, fontSize: 18, italic: true, color: RED });
s.text([
  { text: 'Four colleagues link arms to form one unbroken circle: each figure stands for a part of our end-to-end chain (Plan, Source, Make, Deliver), and only together do they create the shape.', options: { breakLine: true } },
  { text: 'People sit literally at the heart of the design: the ring exists because they hold on to each other, with E2E protected in the centre.', options: { breakLine: true } },
  { text: 'The arrowhead turns the circle into continuous forward motion: one journey that never stops improving.', options: { breakLine: true } },
  { text: 'Built only from circles and arcs in Henkel red tones, the mark stays simple and memorable from a slide footer to an event stage, in colour, white or one colour.' },
], { x: 6.2, y: 2.05, w: 6.5, h: 3.5, fontSize: 15, color: INK, valign: 'top', paraSpaceAfter: 10, objectName: 'Rationale' });
[['Plan', RED], ['Source', INK], ['Make', CORAL], ['Deliver', DEEP]].forEach(([r, c], i) => {
  const x = 6.2 + i * 1.62;
  s.oval(x + 0.21, 6.06, 0.21, c, `Chip ${r}`);
  s.text(r, { x: x + 0.52, y: 5.85, w: 1.0, h: 0.42, fontSize: 13, bold: true, color: INK, valign: 'middle' });
});
s.text('4 people  ·  1 ring  ·  1 direction', { x: 6.2, y: 6.5, w: 6.5, h: 0.35, fontSize: 12, color: GREY, charSpacing: 2 });

// ---------- Slide 3: versatility ----------
s = newSlide(WHITE);
s.text('Works everywhere', { x: 0.6, y: 0.45, w: 8, h: 0.6, fontSize: 30, bold: true, color: INK });
s.text('Presentations · communications · events', { x: 0.6, y: 1.05, w: 8, h: 0.4, fontSize: 15, color: GREY });
s.rect(0.6, 1.8, 5.9, 2.4, LIGHT, 'Tile full colour');
drawMark(s, 1.85, 3.0, 0.72);
wordmark(s, 3.0, 2.55, 3.4, 30, { align: 'left', tagSize: 10, tagSpacing: 2 });
s.rect(6.85, 1.8, 5.9, 2.4, RED, 'Tile reversed');
drawMark(s, 8.1, 3.0, 0.72, { colors: [WHITE, WHITE, WHITE, WHITE], text: WHITE, accent: WHITE });
wordmark(s, 9.25, 2.55, 3.4, 30, { align: 'left', color: WHITE, accent: WHITE, tagColor: WHITE, tagSize: 10, tagSpacing: 2 });
s.rect(0.6, 4.5, 5.9, 2.4, INK, 'Tile dark');
drawMark(s, 1.85, 5.7, 0.72, { colors: [RED, WHITE, WHITE, WHITE], text: WHITE, accent: RED });
wordmark(s, 3.0, 5.25, 3.4, 30, { align: 'left', color: WHITE, tagColor: 'BBBBBB', tagSize: 10, tagSpacing: 2 });
s.rect(6.85, 4.5, 5.9, 2.4, LIGHT, 'Tile sizes');
[[0.58, 8.0], [0.4, 9.45], [0.26, 10.6], [0.16, 11.45]].forEach(([r, x]) => drawMark(s, x, 5.5, r, { center: r > 0.2 }));
s.text('Scales down to an icon: badges, lanyards, slide footers, app tiles', { x: 7.15, y: 6.3, w: 5.4, h: 0.4, fontSize: 12, color: GREY });

// ---------- Logo only (transparent) for SVG/PNG export ----------
const logoOnly = [];
{ const save = svgs.length; s = newSlide(WHITE); drawMark(s, 13.333 / 2, 2.85, 1.75); wordmark(s, 2.67, 5.35, 8, 40);
  logoOnly.push(...svgs.pop().slice(1)); pres._slides.pop(); }

const wrap = (body, vb = `0 0 ${n4(13.333)} ${n4(7.5)}`, w = 1280, h = 720) =>
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${vb}" width="${w}" height="${h}">${body.join('\n')}</svg>`;
fs.mkdirSync('preview', { recursive: true });
svgs.forEach((b, i) => fs.writeFileSync(`preview/slide${i + 1}.svg`, wrap(b)));
fs.writeFileSync('SCR_E2E_Logo.svg', wrap(logoOnly, `${n4(2.2)} ${n4(0.45)} ${n4(8.93)} ${n4(6.6)}`, 857, 634));
pres.writeFile({ fileName: 'SCR_E2E_Logo_Proposal.pptx' }).then(() => console.log('ok', svgs.length));
