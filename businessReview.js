'use strict';

// businessReview.js – VoApps Delivery Intelligence business review slide generator
// Generates a branded .pptx alongside the Excel Delivery Intelligence Report

const pptxgen = require('pptxgenjs');
const path    = require('path');

// ─── Brand palette (from 2026 VoApps master deck theme) ───────────────────────
const NAVY         = '0D053F';   // dk2: deep navy – primary dark
const PINK         = 'FF4B7D';   // accent1: hot pink – primary accent
const PURPLE       = '3F2FB8';   // accent2: purple – secondary accent
const BLUE         = '16509B';   // accent3: blue – tertiary
const CHARCOAL     = '2E2C3E';   // accent4: dark charcoal
const PINK_PALE    = 'FAD6D7';   // accent5: pale pink
const PINK_LIGHT   = 'FF93B1';   // accent6: light pink
const PINK_MED     = 'FF6F97';   // mid pink – card accent variant
const PURPLE_LIGHT = '6558C6';   // lighter purple – card accent variant
const BLUE_LIGHT   = '4473AF';   // lighter blue – card accent variant
const CREAM        = 'FBF7F3';   // lt2: cream – body background
const WHITE        = 'FFFFFF';
const TEXT_DARK    = '0D053F';   // navy for headings on light backgrounds
const TEXT_MID     = '2E2C3E';   // charcoal for body text
const TEXT_SOFT    = '6B6478';   // muted for supporting text

// Semantic colors (kept for data traffic-light logic)
const GREEN        = '1E7E34';
const GREEN_PALE   = 'D4EDDA';
const AMBER        = '856404';
const AMBER_PALE   = 'FFF3CD';
const RED          = 'C0392B';
const RED_PALE     = 'FDECEA';

const SLIDE_W   = 13.33;
const SLIDE_H   = 7.5;
const HEADER_H  = 1.0;
const CONTENT_Y = HEADER_H + 0.18;

// ─── Shape type constant (instance-level in pptxgenjs 4.x) ────────────────────
const RECT = 'rect';

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Header bar: deep navy left-to-right bar with pink logo area on the left.
 * Matches the VoApps brand – navy background, white title text, pink icon block.
 * @param {string} [subheading] - Optional date range shown in PINK_LIGHT below the title.
 * @returns {number} Effective header height (use as base for content Y positioning).
 */
function headerBar(pptx, slide, title, logoPath, subheading) {
  // Header is always HEADER_H – subheading overlays inside (no vertical expansion)
  const totalH = HEADER_H;

  // Full-width navy header
  slide.addShape(RECT, {
    x: 0, y: 0, w: SLIDE_W, h: totalH,
    fill: { color: NAVY }, line: { color: NAVY }
  });

  // Pink icon square on the left – always a perfect square
  const iconBlockW = HEADER_H;
  slide.addShape(RECT, {
    x: 0, y: 0, w: iconBlockW, h: HEADER_H,
    fill: { color: PINK }, line: { color: PINK }
  });

  if (logoPath) {
    const inset = 0.1;
    slide.addImage({
      path: logoPath,
      x: inset,
      y: inset,
      w: iconBlockW - inset * 2,
      h: HEADER_H - inset * 2
    });
  }

  // Title text sits to the right of the icon block
  const titleX = iconBlockW + 0.22;
  const titleW = SLIDE_W - titleX - 0.3;

  // Title covers the full header height, vertically centered
  slide.addText(title, {
    x: titleX, y: 0, w: titleW, h: HEADER_H,
    fontSize: 22, bold: true, color: WHITE,
    fontFace: 'Aktiv Grotesk VF Medium',
    valign: 'middle', align: 'left',
    charSpacing: 0.5
  });

  if (subheading) {
    // Date subtitle overlaid inside the header, close to the title baseline
    slide.addText(subheading, {
      x: titleX, y: 0.66, w: titleW, h: 0.22,
      fontSize: 11, bold: false, color: PINK_LIGHT,
      fontFace: 'Aktiv Grotesk VF Medium',
      valign: 'middle', align: 'left'
    });
  }

  // Thin pink bottom accent line
  slide.addShape(RECT, {
    x: 0, y: totalH - 0.04, w: SLIDE_W, h: 0.04,
    fill: { color: PINK }, line: { color: PINK }
  });

  return totalH;
}

/**
 * Footer bar: cream strip with centered muted text and a thin pink top rule.
 */
function slideFooter(slide, text) {
  const footerH = 0.3;
  const footerY = SLIDE_H - footerH;

  slide.addShape(RECT, {
    x: 0, y: footerY - 0.03, w: SLIDE_W, h: 0.03,
    fill: { color: PINK_PALE }, line: { color: PINK_PALE }
  });

  slide.addText(text || 'VoApps Delivery Intelligence Report', {
    x: 0, y: footerY,
    w: SLIDE_W, h: footerH,
    fontSize: 8, color: TEXT_SOFT, align: 'center',
    fontFace: 'Aktiv Grotesk VF Medium',
    italic: true
  });
}

/**
 * Metric card – cream background, colored top strip, label, large value, subtext.
 * Styled to match VoApps brand card pattern.
 */
function metricBox(slide, x, y, w, h, label, value, subtext, accentColor, valueFontSize) {
  const accent   = accentColor  || PINK;
  const valSize  = valueFontSize || 32;
  const stripH   = 0.06;
  // Value box height: shrinks to leave room for subtext when card is short
  const valBoxH  = subtext ? Math.min(0.68, h - 0.56) : 0.68;

  // Card background – cream
  slide.addShape(RECT, {
    x, y, w, h,
    fill: { color: CREAM },
    line: { color: PINK_PALE, pt: 1 }
  });

  // Colored top strip
  slide.addShape(RECT, {
    x, y, w, h: stripH,
    fill: { color: accent }, line: { color: accent }
  });

  // Label
  slide.addText(label, {
    x: x + 0.16, y: y + stripH + 0.1,
    w: w - 0.32, h: 0.3,
    fontSize: 11, color: TEXT_SOFT,
    bold: false, align: 'left',
    fontFace: 'Aktiv Grotesk VF Medium',
    charSpacing: 1.5
  });

  // Big value
  slide.addText(value, {
    x: x + 0.16, y: y + stripH + 0.38,
    w: w - 0.32, h: valBoxH,
    fontSize: valSize, bold: true, color: NAVY,
    fontFace: 'IvyPresto Text',
    align: 'left', valign: 'top',
    shrinkText: true
  });

  if (subtext) {
    // Anchored after the value box so it never overlaps the number
    slide.addText(subtext, {
      x: x + 0.16, y: y + stripH + 0.38 + valBoxH + 0.08,
      w: w - 0.32, h: 0.3,
      fontSize: 10, color: TEXT_SOFT,
      italic: false, align: 'left',
      fontFace: 'Aktiv Grotesk VF Medium'
    });
  }
}

/**
 * Draw a thermometer shape centred at (cx) with the tube top at tubeTopY.
 *
 * @param {object} slide       - pptxgenjs slide object
 * @param {number} cx          - horizontal centre of the thermometer (inches)
 * @param {number} tubeTopY    - Y of the very top of the tube (inches)
 * @param {number} tubeH       - height of the tube (inches)
 * @param {number} tubeW       - outer width of the tube (inches)
 * @param {number} pct         - fill level 0–100
 * @param {string} fillColor   - hex fill colour (no #)
 * @param {string} title       - label shown below the bulb
 * @param {string} valueTxt    - text shown inside the bulb (e.g. "78%")
 */
function drawThermometer(slide, cx, tubeTopY, tubeH, tubeW, pct, fillColor, title, valueTxt) {
  const pad    = 0.055;              // inner padding on each side
  const innerW = tubeW - pad * 2;
  const tubeX  = cx - tubeW / 2;
  const innerX = tubeX + pad;
  const bulbR  = tubeW * 0.9;       // bulb radius — larger than tube for visual pop

  // ── Outer tube shell (light gray capsule) ────────────────────────────────
  slide.addShape('roundRect', {
    x: tubeX, y: tubeTopY, w: tubeW, h: tubeH,
    fill: { color: 'E4E0EB' },
    line: { color: 'CECCDA', pt: 1 },
    rectRadius: tubeW / 2              // fully rounded ends = capsule
  });

  // ── Coloured fill (bottom-aligned, extends into bulb) ────────────────────
  const clampPct = Math.max(0, Math.min(100, pct));
  const usableH  = tubeH - pad * 2;
  const fillH    = usableH * (clampPct / 100) + pad;  // +pad bridges fill→bulb gap
  if (clampPct > 0) {
    const fillY = tubeTopY + pad + (usableH - usableH * (clampPct / 100));
    slide.addShape('roundRect', {
      x: innerX, y: fillY, w: innerW, h: fillH,
      fill: { color: fillColor },
      line: { color: fillColor },
      rectRadius: innerW / 2
    });
  }

  // ── Bulb ─────────────────────────────────────────────────────────────────
  const bulbCy = tubeTopY + tubeH + bulbR * 0.35;  // overlaps tube bottom a little

  // Outer ring (shell)
  slide.addShape('ellipse', {
    x: cx - bulbR - 0.04, y: bulbCy - bulbR - 0.04,
    w: (bulbR + 0.04) * 2, h: (bulbR + 0.04) * 2,
    fill: { color: 'E4E0EB' },
    line: { color: 'CECCDA', pt: 1 }
  });

  // Coloured fill circle
  slide.addShape('ellipse', {
    x: cx - bulbR, y: bulbCy - bulbR,
    w: bulbR * 2, h: bulbR * 2,
    fill: { color: fillColor },
    line: { color: fillColor }
  });

  // ── Scale ticks on right of tube ─────────────────────────────────────────
  const tickX = tubeX + tubeW + 0.04;
  [0, 25, 50, 75, 100].forEach(m => {
    const isMajor = m % 50 === 0;
    const tickLen = isMajor ? 0.16 : 0.09;
    const tickY   = tubeTopY + pad + usableH * (1 - m / 100) - 0.01;
    slide.addShape(RECT, {
      x: tickX, y: tickY, w: tickLen, h: 0.02,
      fill: { color: isMajor ? 'A09AB0' : 'C0BAD0' },
      line: { color: isMajor ? 'A09AB0' : 'C0BAD0' }
    });
    if (isMajor) {
      slide.addText(`${m}%`, {
        x: tickX + tickLen + 0.04, y: tickY - 0.10, w: 0.42, h: 0.22,
        fontSize: 7, color: 'A09AB0', align: 'left', valign: 'middle',
        fontFace: 'Aktiv Grotesk VF Medium'
      });
    }
  });

  // ── Value inside bulb ────────────────────────────────────────────────────
  slide.addText(valueTxt, {
    x: cx - 0.72, y: bulbCy - 0.22, w: 1.44, h: 0.44,
    fontSize: 13, bold: true, color: WHITE,
    align: 'center', valign: 'middle',
    fontFace: 'Aktiv Grotesk VF Medium'
  });

  // ── Label below bulb ─────────────────────────────────────────────────────
  const labelY = bulbCy + bulbR + 0.16;
  slide.addText(title, {
    x: cx - 1.4, y: labelY, w: 2.8, h: 0.32,
    fontSize: 10, bold: true, color: TEXT_MID,
    align: 'center', charSpacing: 0.3,
    fontFace: 'Aktiv Grotesk VF Medium'
  });
}

// ─── Main export ──────────────────────────────────────────────────────────────

/**
 * Horizontal gauge bar — used for the Clean List metric on the gauge slide.
 * Renders a wide capsule fill-bar with the percentage value centred inside it.
 * @param {object} slide      - pptxgenjs slide
 * @param {number} cx         - horizontal centre of the bar
 * @param {number} midY       - vertical centre of the bar
 * @param {number} barW       - total width of the bar
 * @param {number} pct        - fill percentage (0-100)
 * @param {string} fillColor  - hex fill colour (no #)
 * @param {string} title      - label shown below the bar
 * @param {string} valueTxt   - text shown inside the filled portion
 * @param {string} explanation - one-liner explanation shown below the title
 */
function drawHorizontalGauge(slide, cx, midY, barW, pct, fillColor, title, valueTxt, explanation) {
  const barH  = 0.72;
  const barX  = cx - barW / 2;
  const clamp = Math.max(0, Math.min(100, pct));
  const minFill = barH;  // always at least a circle-width so rounded cap looks right

  // Background track
  slide.addShape('roundRect', {
    x: barX, y: midY - barH / 2, w: barW, h: barH,
    fill: { color: 'E4E0EB' }, line: { color: 'CECCDA', pt: 1 },
    rectRadius: barH / 2
  });

  // Coloured fill
  const fillW = Math.max(minFill, barW * (clamp / 100));
  slide.addShape('roundRect', {
    x: barX, y: midY - barH / 2, w: fillW, h: barH,
    fill: { color: fillColor }, line: { color: fillColor },
    rectRadius: barH / 2
  });

  // Value text centred in the filled segment
  slide.addText(valueTxt, {
    x: barX, y: midY - barH / 2, w: fillW, h: barH,
    fontSize: 20, bold: true, color: 'FFFFFF',
    align: 'center', valign: 'middle',
    fontFace: 'IvyPresto Text'
  });

  // Tick marks at 0%, 50%, 100% on the right-hand side of the track
  [0, 50, 100].forEach(m => {
    const tickX = barX + barW * (m / 100);
    const tickTopY = midY - barH / 2 - 0.14;
    slide.addShape(RECT, { x: tickX - 0.01, y: tickTopY, w: 0.02, h: 0.11,
      fill: { color: 'A09AB0' }, line: { color: 'A09AB0' } });
    slide.addText(`${m}%`, { x: tickX - 0.22, y: tickTopY - 0.18, w: 0.44, h: 0.18,
      fontSize: 7, color: 'A09AB0', align: 'center', fontFace: 'Aktiv Grotesk VF Medium' });
  });

  // Label
  const labelY = midY + barH / 2 + 0.18;
  slide.addText(title, {
    x: cx - 2.0, y: labelY, w: 4.0, h: 0.26,
    fontSize: 10, bold: true, color: TEXT_MID,
    align: 'center', charSpacing: 0.3,
    fontFace: 'Aktiv Grotesk VF Medium'
  });
  if (explanation) {
    slide.addText(explanation, {
      x: cx - 2.2, y: labelY + 0.28, w: 4.4, h: 0.42,
      fontSize: 9, color: '8A8298', italic: true,
      align: 'center', fontFace: 'Aktiv Grotesk VF Medium'
    });
  }
}

/**
 * Draw the Delivery Rate speedometer: three fixed color zones (0-60% navy,
 * 60-70% sky-blue target band, 70-100% light gray) with a PINK needle.
 * If pct < 60 the graphic is suppressed and only text is shown.
 *
 * @param {object} slide  - pptxgenjs slide object
 * @param {number} cx     - horizontal center of the column (inches)
 * @param {number} topY   - top of the chart bounding box (inches)
 * @param {number} size   - diameter of the doughnut bounding box (inches)
 * @param {number} pct    - delivery rate 0-100
 */
function drawDeliverySpeedometer(slide, cx, topY, size, pct) {
  const clamped  = Math.max(0, Math.min(100, pct));
  const LOW = 60, HIGH = 70;
  const C_NAVY  = '0D053F';
  const C_SKY   = '4A9EE0';
  const C_LGRAY = 'C8C4DE';

  // Chart center — needle pivots here
  const circleCy = topY + size / 2;
  const outerR   = size / 2;             // chart radius in inches
  const holeRatio = 0.60;                // holeSize=60 → inner r = 60% of outerR
  const innerR   = outerR * holeRatio;

  // ── Graphic (always shown) ───────────────────────────────────────────
  {
    // Three-zone doughnut: 60 navy | 10 sky | 30 lgray | 100 cream (hidden)
    slide.addChart('doughnut',
      [{ name: 'Delivery', labels: ['low', 'target', 'high', 'hidden'], values: [60, 10, 30, 100] }],
      {
        x: cx - size / 2, y: topY, w: size, h: size,
        holeSize: 60,
        firstSliceAng: 270,
        chartColors: [C_NAVY, C_SKY, C_LGRAY, CREAM],
        dataLabelFontSize: 1,
        showLabel: false, showValue: false, showPercent: false,
        showLegend: false,
        border: { pt: 0, color: CREAM },
        chartBorder: { pt: 0, color: CREAM }
      }
    );

    // ── "60–70% expected" label above the target band ─────────────────
    // Band midpoint at 65% → canvas angle = 180 + 65*1.8 = 297°
    const lblAngRad = (180 + 65 * 1.8) * Math.PI / 180;
    const lblR      = outerR + 0.26;
    const lblCx     = cx       + lblR * Math.cos(lblAngRad);
    const lblCy     = circleCy + lblR * Math.sin(lblAngRad);
    slide.addText('60\u201370% expected', {
      x: lblCx - 0.62, y: lblCy - 0.14,
      w: 1.24, h: 0.30,
      fontSize: 8.5, bold: true, color: C_SKY,
      align: 'center', fontFace: 'Aktiv Grotesk VF Medium'
    });

    // ── Needle ────────────────────────────────────────────────────────
    const needleLen  = innerR * 0.92;
    const nAngRad    = (180 + clamped * 1.8) * Math.PI / 180;
    const ntx        = cx       + needleLen * Math.cos(nAngRad);
    const nty        = circleCy + needleLen * Math.sin(nAngRad);
    const ndx = ntx - cx, ndy = nty - circleCy;
    const adx = Math.abs(ndx), ady = Math.abs(ndy);

    // Line from pivot → tip, accounting for direction with flipH/flipV
    let lx, ly, lw, lh, flipH = false, flipV = false;
    if      (ndx >= 0 && ndy >= 0) { lx = cx;  ly = circleCy; lw = adx; lh = ady; }
    else if (ndx <  0 && ndy >= 0) { lx = ntx; ly = circleCy; lw = adx; lh = ady; flipH = true; }
    else if (ndx >= 0 && ndy <  0) { lx = cx;  ly = nty;      lw = adx; lh = ady; flipV = true; }
    else                            { lx = ntx; ly = nty;      lw = adx; lh = ady; flipH = true; flipV = true; }
    slide.addShape('line', { x: lx, y: ly, w: lw, h: lh, flipH, flipV, line: { color: PINK, pt: 2 } });

    // Pivot dot
    const pivR = 0.09;
    slide.addShape('ellipse', {
      x: cx - pivR, y: circleCy - pivR, w: pivR * 2, h: pivR * 2,
      fill: { color: PINK }, line: { color: PINK }
    });
  }

  // ── Value — pushed down so it sits visibly inside the arc ────────────
  const valY = topY + size * 0.30;
  slide.addText(`${pct.toFixed(1)}%`, {
    x: cx - 1.1, y: valY, w: 2.2, h: 0.5,
    fontSize: 26, bold: true, color: '0D053F',
    align: 'center', fontFace: 'IvyPresto Text'
  });

  // ── Title + explanation — placed at arc center, not below full bounding box ──
  // The doughnut chart h = size but only the top half is visible.
  // Position title just past the midpoint so it sits under the arc with no gap.
  const titY = topY + size * 0.55;
  slide.addText('DELIVERY RATE', {
    x: cx - 1.55, y: titY, w: 3.1, h: 0.30,
    fontSize: 11, bold: true, color: TEXT_MID,
    align: 'center', charSpacing: 0.3, fontFace: 'Aktiv Grotesk VF Medium'
  });
  slide.addText('of all attempts, the % that successfully delivered', {
    x: cx - 1.55, y: titY + 0.34, w: 3.1, h: 0.44,
    fontSize: 10.5, color: '8A8298', italic: true,
    align: 'center', fontFace: 'Aktiv Grotesk VF Medium'
  });
}

/**
 * Draw a half-donut speedometer (top arc visible, bottom hidden in cream).
 * The 60–70% target band is always visible as a pale-green arc segment.
 * Fill color changes to reflect whether the value is below / in / above the band.
 */
function drawSpeedometer(slide, cx, topY, size, pct, title, explanation) {
  const clamped = Math.max(0, Math.min(100, pct));
  const LOW = 60, HIGH = 70;

  const C_FILL_LOW   = PURPLE;     // below target
  const C_FILL_MID   = '2D8A60';   // green — inside target band
  const C_FILL_HIGH  = PINK;       // above target
  const C_TARGET     = 'B8E0CA';   // pale green — unfilled target zone marker
  const C_UNFILLED   = 'DDD9EF';   // light lavender — unfilled arc
  const C_HIDDEN     = CREAM;      // matches slide background — hidden bottom half

  // Visible 100 units (top half arc), hidden 100 units (bottom half, cream)
  const seg_purpleFill   = Math.min(clamped, LOW);
  const seg_greenFill    = Math.max(0, Math.min(clamped, HIGH) - LOW);
  const seg_pinkFill     = Math.max(0, clamped - HIGH);
  const seg_unfilledPre  = Math.max(0, LOW - clamped);
  const seg_targetZone   = (HIGH - LOW) - seg_greenFill;   // remaining pale-green band
  const seg_unfilledPost = Math.max(0, 100 - Math.max(clamped, HIGH));
  const seg_hidden       = 100;

  const rawSegs = [
    { v: seg_purpleFill,   c: C_FILL_LOW  },
    { v: seg_greenFill,    c: C_FILL_MID  },
    { v: seg_pinkFill,     c: C_FILL_HIGH },
    { v: seg_unfilledPre,  c: C_UNFILLED  },
    { v: seg_targetZone,   c: C_TARGET    },
    { v: seg_unfilledPost, c: C_UNFILLED  },
    { v: seg_hidden,       c: C_HIDDEN    },
  ].filter(s => s.v > 0);

  slide.addChart('doughnut',
    [{ name: 'Speedometer', labels: rawSegs.map((_, i) => `s${i}`), values: rawSegs.map(s => s.v) }],
    {
      x: cx - size / 2, y: topY, w: size, h: size,
      holeSize: 62,
      firstSliceAng: 270,           // start at 9 o'clock → arc through 12 → 3 o'clock
      chartColors: rawSegs.map(s => s.c),
      dataLabelFontSize: 1,
      showLabel: false, showValue: false, showPercent: false,
      showLegend: false,
      border: { pt: 0, color: 'FFFFFF' },
      chartBorder: { pt: 0, color: 'FFFFFF' }
    }
  );

  // Value color reflects zone
  const valColor = clamped < LOW ? C_FILL_LOW : (clamped <= HIGH ? C_FILL_MID : C_FILL_HIGH);
  const circCenterY = topY + size / 2;

  // Large value — positioned in the visible upper arc hole
  slide.addText(`${pct.toFixed(1)}%`, {
    x: cx - size / 2, y: topY + size * 0.17, w: size, h: size * 0.37,
    fontSize: 28, bold: true, color: valColor,
    align: 'center', valign: 'middle', fontFace: 'IvyPresto Text'
  });

  // "Target: 60–70%" annotation — just below the arc midpoint, in the cream area
  slide.addText('Target: 60\u201370%', {
    x: cx - 1.1, y: circCenterY + 0.06, w: 2.2, h: 0.22,
    fontSize: 8, bold: false, color: '2D8A60',
    align: 'center', fontFace: 'Aktiv Grotesk VF Medium'
  });

  // Title & explanation stack below in the cream lower half
  slide.addText(title, {
    x: cx - 1.9, y: circCenterY + 0.32, w: 3.8, h: 0.26,
    fontSize: 10, bold: true, color: TEXT_MID,
    align: 'center', charSpacing: 0.3, fontFace: 'Aktiv Grotesk VF Medium'
  });
  if (explanation) {
    slide.addText(explanation, {
      x: cx - 2.0, y: circCenterY + 0.60, w: 4.0, h: 0.44,
      fontSize: 8.5, color: '8A8298', italic: true,
      align: 'center', fontFace: 'Aktiv Grotesk VF Medium'
    });
  }
}

/**
 * Generate a branded business review .pptx alongside the Excel report.
 *
 * @param {Object} stats      - Aggregated stats from generateTrendAnalysis
 * @param {string} outputPath - Destination file path (should end in .pptx)
 * @param {string} logoPath   - Legacy single-logo path (used as fallback for both)
 * @param {string} squareLogo - Square logo for slide headers (pink square icon)
 * @param {string} circleLogo - Circle logo for the title slide
 */
async function generateBusinessReviewSlides(stats, outputPath, logoPath, squareLogo, circleLogo, options = {}) {
  const {
    includeSlideDecayCurve       = false,
    includeSlideReAttemptCadence = true,
    includeSlideOpportunities    = true,
    overviewCards                = null,
    reAttemptData                = null,
    clientPrefix                 = ''
  } = options;

  // ── Helper: strip Excel tab cross-references from slide text, return separately ──
  const TAB_REF_RE = /\s*(?:See\s+[""]?[^."]+[""]?\s+(?:tab\s+)?for[^.]*\.|The\s+[^.]+?\s+tab\s+has[^.]*\.)/gi;
  function separateTabRefs(text) {
    const notesList = [...text.matchAll(new RegExp(TAB_REF_RE.source, 'gi'))].map(m => m[0].trim());
    const visible = text.replace(new RegExp(TAB_REF_RE.source, 'gi'), '').trim();
    return { visible, notes: notesList };
  }

  // Resolve which logo to use for each context
  const headerLogo = squareLogo || logoPath;
  const titleLogo  = circleLogo || logoPath;
  const {
    uniqueNumbers,
    totalAttempts,
    totalSuccess,
    overallSuccessRate,
    listGrade,
    healthyCount,
    toxicCount,
    neverDeliveredCount,
    suppressionCandidateCount,
    avgVariability,
    decayCurve,
    cadence,
    actions,
    bestNextAction,
    agentHoursSaved,
    staleWarmCount,
    impliedRemovedCount = 0,
    minDate,
    maxDate,
    accountIds
  } = stats;

  // Implied callbacks: numbers delivered then not re-attempted within one cadence window.
  // These are the most likely source of inbound callbacks after a campaign.
  const impliedCallbackRate = totalSuccess > 0
    ? (impliedRemovedCount / totalSuccess * 100)
    : 0;

  const pptx = new pptxgen();
  pptx.layout  = 'LAYOUT_WIDE';
  pptx.author  = 'VoApps Tools';
  pptx.subject = 'Delivery Intelligence Report';
  pptx.title   = 'VoApps Delivery Intelligence Report';

  const fmtDate = d => d
    ? d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }).toUpperCase()
    : '';
  const dateRangeStr = (minDate && maxDate)
    ? `${fmtDate(minDate)} – ${fmtDate(maxDate)}`
    : 'Date range not available';

  const daySpan = (minDate && maxDate)
    ? Math.round((maxDate.getTime() - minDate.getTime()) / (1000 * 60 * 60 * 24))
    : 0;

  const acctList = (accountIds && accountIds.length > 0)
    ? accountIds.join(', ')
    : '';

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 1 – Title
  // Full dark navy background with centered content, pink accent bar, logo
  // ────────────────────────────────────────────────────────────────────────────
  const s1 = pptx.addSlide();

  // Full navy background
  s1.addShape(RECT, {
    x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
    fill: { color: NAVY }, line: { color: NAVY }
  });

  // Decorative pink horizontal bar near top-third
  s1.addShape(RECT, {
    x: 0, y: 2.7, w: SLIDE_W, h: 0.05,
    fill: { color: PINK }, line: { color: PINK }
  });

  // Pink accent bottom strip
  s1.addShape(RECT, {
    x: 0, y: SLIDE_H - 0.06, w: SLIDE_W, h: 0.06,
    fill: { color: PINK }, line: { color: PINK }
  });

  // Logo centered (circle logo on dark navy background)
  if (titleLogo) {
    const logoSize = 1.6;
    s1.addImage({
      path: titleLogo,
      x: SLIDE_W / 2 - logoSize / 2,
      y: 0.7,
      w: logoSize,
      h: logoSize
    });
  }

  // Main title
  s1.addText('Delivery Intelligence Report', {
    x: 0.5, y: 2.85, w: SLIDE_W - 1.0, h: 1.0,
    fontSize: 40, bold: true, color: WHITE,
    fontFace: 'IvyPresto Text',
    align: 'center', valign: 'middle'
  });

  // Subtitle / date range
  s1.addText(dateRangeStr, {
    x: 0.5, y: 3.92, w: SLIDE_W - 1.0, h: 0.52,
    fontSize: 17, color: PINK_LIGHT,
    fontFace: 'Aktiv Grotesk VF Medium',
    align: 'center'
  });

  // Filename prefix (client name) just above Account line
  if (clientPrefix) {
    s1.addText(clientPrefix, {
      x: 0.5, y: 4.44, w: SLIDE_W - 1.0, h: 0.30,
      fontSize: 13, bold: true, color: WHITE,
      fontFace: 'Aktiv Grotesk VF Medium',
      align: 'center'
    });
  }

  if (acctList) {
    s1.addText(`Account${accountIds.length > 1 ? 's' : ''}: ${acctList}`, {
      x: 0.5, y: clientPrefix ? 4.76 : 4.5, w: SLIDE_W - 1.0, h: 0.5,
      fontSize: 12, color: PINK_PALE,
      fontFace: 'Aktiv Grotesk VF Medium',
      align: 'center', italic: true,
      shrinkText: true
    });
  }

  s1.addText('Generated by VoApps Tools', {
    x: 0, y: SLIDE_H - 0.95, w: SLIDE_W, h: 0.38,
    fontSize: 10, color: '3D2E62',  // dim purple – visible but not a focus element
    fontFace: 'Aktiv Grotesk VF Medium',
    align: 'center', italic: true
  });

  s1.addNotes([
    'TITLE SLIDE — Speaker notes',
    '',
    'This Delivery Intelligence Report was generated automatically by VoApps Tools from raw campaign data.',
    'Date range: ' + dateRangeStr,
    (acctList ? 'Accounts: ' + acctList : ''),
    '',
    'Use this slide to orient the audience: what time period are we reviewing, which accounts are included, and what is the purpose of this review.',
    'The report covers three dimensions: delivery efficiency (how well DirectDrop Voicemail drops are connecting), list health (which numbers are still productive), and re-attempt strategy (how much incremental value multi-touch delivers).'
  ].filter(Boolean).join('\n'));

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 2 – Delivery Performance Snapshot
  // Large speedometer (left 2/3) + Agent Hours Saved (right 1/3)
  // Fixed bottom stats: Unique Phone Numbers | Successful Deliveries | Avg Attempts/Number
  // ────────────────────────────────────────────────────────────────────────────
  const sGauge = pptx.addSlide();
  sGauge.background = { color: CREAM };
  const sGaugeHdrH = headerBar(pptx, sGauge, 'Delivery Performance Snapshot', headerLogo, dateRangeStr);

  const healthyPct = uniqueNumbers > 0 ? (healthyCount / uniqueNumbers) * 100 : 0;

  // Layout: divider at true center; each zone is exactly half the slide
  const dividerX  = SLIDE_W / 2;                // 6.665" — true center
  const spdCx     = dividerX / 2;               // 3.33" — center of left half
  const spdSize   = 3.2;                         // compact — avoids gap from invisible bottom half
  const ahCardW   = 4.40;                        // card width (fits within right half with margins)
  const ahCardH   = 2.46;

  // Vertically center the content block in the body area (between header and stats bar)
  // Speedometer visible height = size*0.55 (title sits at midpoint) + 0.34 + 0.44 (text stack)
  const spdVisualH = spdSize * 0.55 + 0.78;     // 2.54" — arc + DELIVERY RATE + explanation
  const contentH   = Math.max(spdVisualH, ahCardH + 0.10); // taller of the two zones
  const statY      = SLIDE_H - 1.40;            // bottom stats bar top edge
  const thTopY     = sGaugeHdrH + Math.max(0.30, (statY - 0.10 - sGaugeHdrH - contentH) / 2);

  // ── 1. Delivery Rate — speedometer centered in left half ─────────────────
  drawDeliverySpeedometer(sGauge, spdCx, thTopY, spdSize, overallSuccessRate);

  // ── 2. Agent Hours Saved — card centered in right half ───────────────────
  const ahCardX  = dividerX + (dividerX - ahCardW) / 2;  // centered in right half
  const ahCardY  = thTopY + 0.10;

  // Thin vertical separator at true center — 70% of content area height
  const divH = (SLIDE_H - sGaugeHdrH - 1.55) * 0.70;
  sGauge.addShape(RECT, {
    x: dividerX, y: sGaugeHdrH + 0.30,
    w: 0.02, h: divH,
    fill: { color: 'E0DCEA' }, line: { color: 'E0DCEA' }
  });

  // Card: light pink background + pink border
  sGauge.addShape(RECT, {
    x: ahCardX, y: ahCardY, w: ahCardW, h: ahCardH,
    fill: { color: 'FEF0F3' }, line: { color: PINK_PALE, pt: 0.75 }
  });
  // Pink accent bar at top of card
  sGauge.addShape(RECT, { x: ahCardX, y: ahCardY, w: ahCardW, h: 0.06, fill: { color: PINK }, line: { color: PINK } });

  sGauge.addText('AGENT HOURS SAVED (EST.)', {
    x: ahCardX + 0.16, y: ahCardY + 0.16, w: ahCardW - 0.32, h: 0.28,
    fontSize: 11, bold: true, color: TEXT_SOFT,
    align: 'center', charSpacing: 1, fontFace: 'Aktiv Grotesk VF Medium'
  });
  sGauge.addText(agentHoursSaved > 0 ? agentHoursSaved.toLocaleString() : '—', {
    x: ahCardX + 0.16, y: ahCardY + 0.52, w: ahCardW - 0.32, h: 1.10,
    fontSize: 52, bold: true, color: NAVY,
    align: 'center', valign: 'middle', fontFace: 'IvyPresto Text', shrinkText: true
  });
  sGauge.addText(
    agentHoursSaved > 0
      ? `Based on ${totalSuccess.toLocaleString()} deliveries × 3 min avg handle time — capacity freed for higher-value work`
      : 'Enable agent hours calculation by confirming avg handle time in settings',
    {
      x: ahCardX + 0.16, y: ahCardY + 1.72, w: ahCardW - 0.32, h: 0.62,
      fontSize: 11, color: '8A8298', italic: true,
      align: 'center', fontFace: 'Aktiv Grotesk VF Medium'
    }
  );

  // ── Bottom stats row: fixed — Unique Phone Numbers | Successful Deliveries | Avg Attempts ──
  const gcX   = 0.5;
  const gcW   = SLIDE_W - 1.0;
  // statY already declared above for vertical centering calculation

  sGauge.addShape(RECT, {
    x: gcX, y: statY - 0.04, w: gcW, h: 0.02,
    fill: { color: 'E0DCEA' }, line: { color: 'E0DCEA' }
  });

  const avgAttemptsVal = uniqueNumbers > 0 ? (totalAttempts / uniqueNumbers).toFixed(1) : '—';
  const gaugeStats = [
    { value: uniqueNumbers.toLocaleString(),   label: 'UNIQUE PHONE NUMBERS'    },
    { value: totalSuccess.toLocaleString(),    label: 'SUCCESSFUL DELIVERIES'   },
    { value: avgAttemptsVal,                   label: 'AVG. ATTEMPTS PER NUMBER'}
  ];
  const statColW = gcW / 3;
  gaugeStats.forEach((st, i) => {
    const sx = gcX + i * statColW;
    if (i > 0) {
      sGauge.addShape(RECT, {
        x: sx, y: statY + 0.06, w: 0.02, h: 0.68,
        fill: { color: 'DEDEDE' }, line: { color: 'DEDEDE' }
      });
    }
    sGauge.addText(st.value, {
      x: sx, y: statY, w: statColW, h: 0.58,
      fontSize: 26, bold: true, color: NAVY,
      fontFace: 'IvyPresto Text', align: 'center'
    });
    sGauge.addText(st.label, {
      x: sx, y: statY + 0.55, w: statColW, h: 0.26,
      fontSize: 9.5, color: TEXT_SOFT,
      fontFace: 'Aktiv Grotesk VF Medium', align: 'center', charSpacing: 1
    });
  });

  sGauge.addNotes([
    'DELIVERY PERFORMANCE SNAPSHOT — Speaker notes',
    '',
    'DELIVERY RATE (' + overallSuccessRate.toFixed(1) + '%)',
    'This is the headline efficiency metric: the share of DirectDrop Voicemail drops that successfully landed in a live voicemail inbox.',
    'The speedometer shows where the rate falls relative to the 60–70% expected range. Rates above 70% reflect excellent list health and carrier connectivity. Rates in the 60–70% band are normal and healthy.',
    overallSuccessRate >= 70
      ? 'Great news: this rate is above the expected range — the list is performing exceptionally well.'
      : overallSuccessRate >= 60
        ? 'This rate is right in the expected range — the campaign is running efficiently.'
        : 'Opportunity: a few list hygiene steps (removing non-deliverable numbers) can meaningfully move this rate into the expected range.',
    '',
    'AGENT HOURS SAVED (' + (agentHoursSaved > 0 ? agentHoursSaved.toLocaleString() + ' hrs' : 'not calculated') + ')',
    'Estimated based on ' + totalSuccess.toLocaleString() + ' successful deliveries × 3 minutes average handle time.',
    'This is the operational capacity DirectDrop Voicemail returned to the business — time agents did not have to spend leaving manual voicemails.',
    'Use this figure to anchor ROI conversations: the hours freed represent real cost savings or capacity available for higher-value outreach.',
    '',
    'BOTTOM STATS',
    '• Unique Phone Numbers (' + uniqueNumbers.toLocaleString() + '): the size of the list being worked.',
    '• Successful Deliveries (' + totalSuccess.toLocaleString() + '): voicemails that reached a live inbox.',
    '• Avg. Attempts Per Number (' + avgAttemptsVal + '): higher values indicate a persistent multi-touch strategy.',
  ].join('\n'));
  slideFooter(sGauge);

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 3 – High-Level Overview
  // 2-row × 3-col metric cards on cream background
  // ────────────────────────────────────────────────────────────────────────────
  const s2 = pptx.addSlide();
  s2.background = { color: CREAM };
  const s2HdrH = headerBar(pptx, s2, 'High-Level Overview', headerLogo, dateRangeStr);

  const bW   = 4.0;
  const bH   = 1.62;
  const bGap = 0.16;
  const row1Y = s2HdrH + 0.5;   // 1.5 – matches user-edited layout
  const row2Y = row1Y + bH + bGap;
  const c1 = 0.32;
  const c2 = c1 + bW + bGap;
  const c3 = c2 + bW + bGap;

  const successAccent = overallSuccessRate >= 75 ? BLUE
    : overallSuccessRate >= 50 ? PURPLE_LIGHT
    : CHARCOAL;
  // healthyPct already declared above for the thermometer section
  const healthyAccent = healthyPct >= 80 ? BLUE : healthyPct >= 60 ? BLUE_LIGHT : PURPLE_LIGHT;

  // ── Card registry – limited to the 4 selectable overview cards ──────────
  // firstAttemptSuccessRate uses decayCurve[0] if available
  const firstAttemptRate = (decayCurve && decayCurve.length > 0)
    ? `${(decayCurve[0].probability * 100).toFixed(1)}%`
    : `${overallSuccessRate.toFixed(1)}%`;
  const firstAttemptPct = (decayCurve && decayCurve.length > 0)
    ? decayCurve[0].probability * 100
    : overallSuccessRate;

  const ALL_CARDS = {
    firstAttemptSuccessRate: {
      label: 'FIRST ATTEMPT SUCCESS RATE',
      value: firstAttemptRate,
      sub: 'Success rate on the very first delivery attempt to each number — baseline list quality before multi-touch selection bias',
      accent: BLUE_LIGHT,
      barPct: firstAttemptPct
    },
    avgAttemptsPerNumber: {
      label: 'AVG. ATTEMPTS PER NUMBER',
      value: uniqueNumbers > 0 ? (totalAttempts / uniqueNumbers).toFixed(1) : '—',
      sub: 'Average total delivery attempts per unique phone number',
      accent: PURPLE_LIGHT
    },
    impliedCallbackOppty: {
      label: 'IMPLIED CALLBACK RATE',
      value: `${impliedCallbackRate.toFixed(1)}%`,
      sub: `${impliedRemovedCount.toLocaleString()} delivered numbers appear removed — each is a potential inbound callback`,
      accent: PINK_MED
    },
    dateSpan: {
      label: 'DATE SPAN',
      value: daySpan > 0 ? `${daySpan} Days` : dateRangeStr,
      sub: daySpan > 0 ? dateRangeStr : '',
      accent: PINK_LIGHT,
      fontSize: 28
    }
  };

  // All 4 are available; user selection filters which ones appear.
  const ALLOWED_CARD_KEYS = ['firstAttemptSuccessRate', 'avgAttemptsPerNumber', 'impliedCallbackOppty', 'dateSpan'];
  const DEFAULT_CARDS = ALLOWED_CARD_KEYS;
  const cardKeys = (Array.isArray(overviewCards) && overviewCards.length > 0)
    ? overviewCards.filter(k => ALLOWED_CARD_KEYS.includes(k)).slice(0, 4)
    : DEFAULT_CARDS;

  // Always 4-per-page (2×2) — max 4 cards, one slide only
  const p4W   = 5.8;
  const p4H   = 1.90;
  const p4Gap = 0.26;
  const p4c1  = (SLIDE_W - 2 * p4W - p4Gap) / 2;
  const p4c2  = p4c1 + p4W + p4Gap;
  const p4r1  = s2HdrH + 0.55;
  const p4r2  = p4r1 + p4H + p4Gap;
  const pos4  = [
    [p4c1, p4r1], [p4c2, p4r1],
    [p4c1, p4r2], [p4c2, p4r2]
  ];

  // Helper: render cards onto a slide — bar-style for firstAttemptSuccessRate
  function renderCardPage(slide, pageKeys) {
    pageKeys.forEach((key, i) => {
      const card = ALL_CARDS[key];
      const pos  = pos4[i];
      if (!card || !pos) return;
      const [cx, cy] = pos;
      const cW = p4W, cH = p4H;

      if (key === 'firstAttemptSuccessRate') {
        // Bar-style card (mirrors overallSuccessRate treatment)
        slide.addShape(RECT, { x: cx, y: cy, w: cW, h: cH, fill: { color: CREAM }, line: { color: PINK_PALE, pt: 1 } });
        slide.addShape(RECT, { x: cx, y: cy, w: cW, h: 0.06, fill: { color: card.accent }, line: { color: card.accent } });
        slide.addText(card.label, { x: cx + 0.16, y: cy + 0.14, w: cW - 0.32, h: 0.26,
          fontSize: 10, color: TEXT_SOFT, fontFace: 'Aktiv Grotesk VF Medium', charSpacing: 1.5 });
        slide.addText(card.value, { x: cx + 0.16, y: cy + 0.40, w: cW - 0.32, h: 0.62,
          fontSize: 34, bold: true, color: NAVY, fontFace: 'IvyPresto Text', align: 'left', valign: 'top', shrinkText: true });
        const pbX = cx + 0.16, pbY = cy + 1.10, pbW2 = cW - 0.32, pbH = 0.14;
        slide.addShape(RECT, { x: pbX, y: pbY, w: pbW2, h: pbH, fill: { color: 'DDD9EF' }, line: { color: 'DDD9EF' } });
        const fillW = Math.max(0.04, pbW2 * (card.barPct / 100));
        slide.addShape(RECT, { x: pbX, y: pbY, w: fillW, h: pbH, fill: { color: card.accent }, line: { color: card.accent } });
        slide.addText(card.sub, { x: cx + 0.16, y: pbY + 0.20, w: cW - 0.32, h: 0.52,
          fontSize: 9.5, color: TEXT_SOFT, fontFace: 'Aktiv Grotesk VF Medium' });
      } else {
        metricBox(slide, cx, cy, cW, cH, card.label, card.value, card.sub, card.accent, card.fontSize || 38);
      }
    });
  }

  // Single overview slide — no s2b
  renderCardPage(s2, cardKeys);

  // ── Single-touch callout — always on s2, below the 2×2 grid ─────────────
  const STRIP_W    = 2 * p4W + p4Gap;
  const STRIP_X    = p4c1;
  const stripBaseY = p4r2 + p4H + 0.24;

  const cadenceTotalNumbers = (cadence.cadenceSingleTouch || 0) + (cadence.cadenceMultiTouchCount || 0);
  if (cadenceTotalNumbers > 0) {
    const stCount = cadence.cadenceSingleTouch;
    const stPct   = (stCount / cadenceTotalNumbers * 100).toFixed(1);
    const staleNote = staleWarmCount > 0
      ? ` ${staleWarmCount.toLocaleString()} confirmed-reachable numbers haven\u2019t been contacted in 30+ days \u2014 ready for immediate re-engagement.`
      : '';
    const impliedNote = impliedRemovedCount > 0
      ? ` ~${impliedRemovedCount.toLocaleString()} numbers appear removed after their last successful delivery \u2014 a potential inbound callback opportunity.`
      : '';

    const calloutH  = 1.06;
    const bigNumW   = 1.60;
    const textX     = STRIP_X + bigNumW + 0.10;
    const textW     = STRIP_W - bigNumW - 0.14;

    // Card background — navy with pink top bar for energy
    s2.addShape(RECT, { x: STRIP_X, y: stripBaseY, w: STRIP_W, h: calloutH,
      fill: { color: '0D053F' }, line: { color: '0D053F' } });
    s2.addShape(RECT, { x: STRIP_X, y: stripBaseY, w: STRIP_W, h: 0.05,
      fill: { color: PINK }, line: { color: PINK } });

    // Large % on the left — visually dominant anchor
    s2.addText(`${stPct}%`, {
      x: STRIP_X, y: stripBaseY + 0.05, w: bigNumW + 0.02, h: calloutH - 0.06,
      fontSize: 32, bold: true, color: PINK,
      fontFace: 'IvyPresto Text', align: 'center', valign: 'middle', shrinkText: true
    });

    // Thin vertical divider
    s2.addShape(RECT, { x: STRIP_X + bigNumW + 0.02, y: stripBaseY + 0.14, w: 0.02, h: calloutH - 0.28,
      fill: { color: 'FFFFFF' }, line: { color: 'FFFFFF' } });
    // Opacity effect via fill alpha isn't directly supported; use a light hex instead
    s2.addShape(RECT, { x: STRIP_X + bigNumW + 0.02, y: stripBaseY + 0.14, w: 0.02, h: calloutH - 0.28,
      fill: { color: '3D3070' }, line: { color: '3D3070' } });

    // Headline
    s2.addText(
      `of numbers (${stCount.toLocaleString()}) received only one attempt this period`,
      { x: textX, y: stripBaseY + 0.10, w: textW, h: 0.32,
        fontSize: 11.5, bold: true, color: WHITE,
        fontFace: 'Aktiv Grotesk VF Medium', valign: 'middle' }
    );
    // Body
    s2.addText(
      `Consumers often need 2\u20133 touches before taking action. A follow-up campaign at a 3\u201310 day interval can produce meaningful incremental results from this same list.${staleNote}${impliedNote}`,
      { x: textX, y: stripBaseY + 0.46, w: textW, h: 0.50,
        fontSize: 9.5, color: 'C8C4DE', italic: true,
        fontFace: 'Aktiv Grotesk VF Medium', valign: 'top' }
    );
  }

  s2.addNotes([
    'HIGH-LEVEL OVERVIEW — Speaker notes',
    '',
    'This slide shows 4 key performance indicators for the period. Use it to anchor the conversation before deeper analysis.',
    '',
    'FIRST ATTEMPT SUCCESS RATE (' + firstAttemptRate + ')',
    'The percentage of numbers that received a successful delivery on the very first attempt — a clean measure of list quality before any multi-touch selection bias.',
    'Comparing this to the overall delivery rate reveals how much incremental value the multi-touch strategy is generating.',
    '',
    'AVG. ATTEMPTS PER NUMBER (' + (uniqueNumbers > 0 ? (totalAttempts / uniqueNumbers).toFixed(1) : '—') + ')',
    'Higher values indicate a persistent multi-touch campaign. Pair with the cadence slide to assess timing strategy and diminishing returns.',
    '',
    'IMPLIED CALLBACK RATE (' + impliedCallbackRate.toFixed(1) + '%)',
    impliedRemovedCount.toLocaleString() + ' delivered numbers appear to have been removed from the list after a successful delivery — consistent with a contact taking inbound action (calling back, submitting a form, making a payment).',
    'This is a compelling ROI signal: it suggests the campaign is generating conversions beyond what outbound metrics alone can capture.',
    '',
    'DATE SPAN',
    'Important context for normalising attempt counts. A longer window with the same attempt count implies lower frequency — useful for pacing conversations.',
    '',
    'SINGLE-TOUCH OPPORTUNITY',
    cadenceTotalNumbers > 0
      ? ((cadence.cadenceSingleTouch / cadenceTotalNumbers * 100).toFixed(1)) + '% of numbers (' + (cadence.cadenceSingleTouch || 0).toLocaleString() + ') received only one attempt this period. A follow-up campaign targeting this group can generate meaningful incremental deliveries with minimal additional list cost.'
      : 'No cadence data available for single-touch analysis.'
  ].join('\n'));
  slideFooter(s2);

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 3 – Success Probability by Attempt (optional)
  // ────────────────────────────────────────────────────────────────────────────
  if (includeSlideDecayCurve) {
  const s3 = pptx.addSlide();
  s3.background = { color: CREAM };
  const s3HdrH = headerBar(pptx, s3, 'Success Probability by Attempt', headerLogo, dateRangeStr);
  const s3CY   = s3HdrH + 0.40;

  s3.addText(
    'Each row represents all numbers at that attempt count. As attempt index rises, the pool shifts toward harder-to-reach numbers – success probability naturally declines.',
    {
      x: 0.4, y: s3CY + 0.08,
      w: SLIDE_W - 0.8, h: 0.4,
      fontSize: 9.5, color: TEXT_SOFT, italic: true, align: 'center',
      fontFace: 'Aktiv Grotesk VF Medium'
    }
  );

  const headerCellOpts = { bold: true, color: WHITE, fill: NAVY, align: 'center', fontFace: 'Aktiv Grotesk VF Medium', fontSize: 10.5 };

  const tblRows = [
    [
      { text: 'Attempt',        options: { ...headerCellOpts, align: 'center' } },
      { text: 'Success Rate',   options: { ...headerCellOpts } },
      { text: 'Successful',     options: { ...headerCellOpts } },
      { text: 'Total Attempts', options: { ...headerCellOpts } },
      { text: 'Insight',        options: { ...headerCellOpts } }
    ],
    ...decayCurve.slice(0, 8).map((dc, idx) => {
      const pct = dc.probability * 100;
      const isEven   = idx % 2 === 0;
      const rowBg    = isEven ? WHITE : 'F5F2EF';
      const rowFill  = pct >= 50 ? 'EBF7EE' : pct >= 25 ? 'FFF8E1' : pct >= 15 ? AMBER_PALE : RED_PALE;
      const txtColor = pct >= 50 ? GREEN    : pct >= 25 ? AMBER    : pct >= 15 ? AMBER      : RED;
      const insight  = pct >= 50 ? 'Good – Continue'
        : pct >= 25 ? 'Declining – Monitor'
        : pct >= 15 ? 'Low – Review'
        : '';

      const cell = (txt, fill, opts) => ({ text: txt, options: { color: TEXT_MID, fill: fill || rowBg, align: 'center', fontFace: 'Aktiv Grotesk VF Medium', fontSize: 10.5, ...opts } });

      return [
        cell(String(dc.attemptIndex)),
        { text: `${pct.toFixed(1)}%`, options: { bold: true, color: txtColor, fill: rowFill, align: 'center', fontFace: 'Aktiv Grotesk VF Medium', fontSize: 10.5 } },
        cell(dc.successful.toLocaleString()),
        cell(dc.total.toLocaleString()),
        { text: insight, options: { color: txtColor, fill: rowFill, align: 'center', fontFace: 'Aktiv Grotesk VF Medium', fontSize: 10.5 } }
      ];
    })
  ];

  {
    const s3TblW  = SLIDE_W - 2.8;
    const s3TblRH = 0.43;
    const s3TblH  = tblRows.length * s3TblRH;
    const s3BodyT = s3CY + 0.58;
    const s3BodyB = SLIDE_H - 0.35;
    const s3TblY  = s3BodyT + Math.max(0, (s3BodyB - s3BodyT - s3TblH) / 2);
    s3.addTable(tblRows, {
      x: (SLIDE_W - s3TblW) / 2, y: s3TblY,
      w: s3TblW,
      fontSize: 10.5,
      rowH: s3TblRH,
      border: { type: 'solid', color: PINK_PALE, pt: 0.75 }
    });
  }

  s3.addNotes([
    'SUCCESS PROBABILITY BY ATTEMPT — Speaker notes',
    '',
    'HOW TO READ THIS TABLE:',
    'Each row represents all phone numbers in the dataset at a specific attempt count. As attempt index increases, the pool naturally shifts toward harder-to-reach numbers — so declining success rates are expected and normal.',
    '',
    'KEY TALKING POINTS:',
    '• Attempt 1 is always the cleanest signal — it reflects the baseline reachability of the list without any selection bias.',
    '• Where does the rate drop below 25%? That is typically the point of diminishing returns — continued attempts produce fewer and fewer new deliveries per drop.',
    '• "Declining – Monitor" rows (25–49%) are still worth retrying if the cadence interval is appropriate (3–10 days).',
    '• Rows marked "Low – Review" or with empty insight labels signal that this portion of the list may benefit from suppression.',
    '',
    'EXCEL REFERENCE:',
    'Numbers with 4–6+ consecutive failures and low success rates are listed in the Suppression Candidates tab of the Delivery Intelligence Excel report.',
    'The Re-Attempt Summary and Outcome Transition Matrix tabs (if included) provide additional granularity on retry ROI.'
  ].join('\n'));
  slideFooter(s3);
  } // end includeSlideDecayCurve

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 4 – Delivery Re-Attempt Cadence (optional)
  // ────────────────────────────────────────────────────────────────────────────
  if (includeSlideReAttemptCadence) {
  const s4 = pptx.addSlide();
  s4.background = { color: CREAM };
  const s4HdrH = headerBar(pptx, s4, 'Delivery Re-Attempt Cadence', headerLogo, dateRangeStr);
  const s4CY   = s4HdrH + 0.40;

  const totalMulti = cadence.cadenceMultiTouchCount || 1;
  const s4TotalNumbers = (cadence.cadenceSingleTouch || 0) + (cadence.cadenceMultiTouchCount || 0);
  const s4SinglePct = s4TotalNumbers > 0 ? (cadence.cadenceSingleTouch / s4TotalNumbers * 100).toFixed(1) : '0.0';
  const s4MultiPct  = s4TotalNumbers > 0 ? (cadence.cadenceMultiTouchCount / s4TotalNumbers * 100).toFixed(1) : '0.0';
  s4.addText(
    `${cadence.cadenceMultiTouchCount.toLocaleString()} numbers (${s4MultiPct}%) had 2+ delivery attempts – across all result types (successfully delivered, unsuccessful, voicemail not setup, voicemail full, and not in service). ${cadence.cadenceSingleTouch.toLocaleString()} numbers (${s4SinglePct}%) were contacted only once – each a potential opportunity for an additional touch. Breakdown by median interval between consecutive attempts (Same-day reflects any pair on the same calendar date):`,
    {
      x: 1.8, y: s4CY + 0.08,
      w: SLIDE_W - 3.6, h: 0.44,
      fontSize: 9.5, color: TEXT_SOFT, italic: true, align: 'center',
      fontFace: 'Aktiv Grotesk VF Medium'
    }
  );

  const cadenceRows = [
    { label: 'Same-day re-attempt (same calendar date)', count: cadence.cadenceBucket_sameDay,   ideal: false, warn: true,  note: false, long: false },
    { label: '1–2 days',                       count: cadence.cadenceBucket_1to2,      ideal: false, warn: true,  note: false, long: false },
    { label: '3–5 days',                       count: cadence.cadenceBucket_3to5,      ideal: true,  warn: false, note: false, long: false },
    { label: '6–10 days',                      count: cadence.cadenceBucket_6to10,     ideal: true,  warn: false, note: false, long: false },
    { label: '11–15 days',                     count: cadence.cadenceBucket_11to15,    ideal: false, warn: false, note: true,  long: false },
    { label: '16–30 days',                     count: cadence.cadenceBucket_16to30,    ideal: false, warn: false, note: false, long: true  },
    { label: '30+ days',                       count: cadence.cadenceBucket_over30,    ideal: false, warn: false, note: false, long: true  }
  ];

  const cadHdrOpts = { bold: true, color: WHITE, fill: NAVY, fontFace: 'Aktiv Grotesk VF Medium', fontSize: 10.5 };

  const cTbl = [
    [
      { text: 'Re-Attempt Interval',  options: { ...cadHdrOpts, align: 'left' } },
      { text: 'Numbers',              options: { ...cadHdrOpts, align: 'center' } },
      { text: '% of Re-Attempted',    options: { ...cadHdrOpts, align: 'center' } },
      { text: '',                     options: { ...cadHdrOpts, align: 'center' } }
    ],
    ...cadenceRows.filter(r => r.count > 0 || r.ideal).map((r, idx) => {
      const rowBg = idx % 2 === 0 ? WHITE : 'F5F2EF';
      const fill  = r.ideal ? GREEN_PALE : r.warn ? RED_PALE : r.note ? AMBER_PALE : r.long ? RED_PALE : rowBg;
      const color = r.ideal ? GREEN      : r.warn ? RED      : r.note ? AMBER      : r.long ? RED      : TEXT_MID;
      const badge = r.ideal ? '✓ Ideal' : r.warn ? '⚠ Too soon' : r.note ? '↑ Longer than ideal' : r.long ? '↑ Consider shorter' : '';
      const badgeColor = r.ideal ? GREEN : r.warn ? RED : r.note ? AMBER : r.long ? RED : TEXT_MID;
      return [
        { text: r.label,                                         options: { color, fill, bold: r.ideal, align: 'left',  fontFace: 'Aktiv Grotesk VF Medium', fontSize: 10.5 } },
        { text: r.count.toLocaleString(),                        options: { color, fill, bold: r.ideal, align: 'center', fontFace: 'Aktiv Grotesk VF Medium', fontSize: 10.5 } },
        { text: `${(r.count / totalMulti * 100).toFixed(1)}%`,  options: { color, fill, bold: r.ideal, align: 'center', fontFace: 'Aktiv Grotesk VF Medium', fontSize: 10.5 } },
        { text: badge,                                            options: { color: badgeColor, fill, bold: r.ideal || r.warn, align: 'center', fontFace: 'Aktiv Grotesk VF Medium', fontSize: 10.5 } }
      ];
    })
  ];

  {
    const s4TblW  = SLIDE_W - 3.6;
    const s4TblRH = 0.44;
    const s4TblH  = cTbl.length * s4TblRH;
    const s4BodyT = s4CY + 0.62;
    const s4BodyB = SLIDE_H - 1.72; // leave room for insight strip (starts at SLIDE_H - 1.62) + gap
    const s4TblY  = s4BodyT + Math.max(0, (s4BodyB - s4BodyT - s4TblH) / 2);
    s4.addTable(cTbl, {
      x: (SLIDE_W - s4TblW) / 2, y: s4TblY,
      w: s4TblW,
      fontSize: 10.5, rowH: s4TblRH,
      border: { type: 'solid', color: PINK_PALE, pt: 0.75 }
    });
  }

  // ── Multi-touch value insight strip ─────────────────────────────────────────
  {
    const insY = SLIDE_H - 1.62;
    const insX = 1.8;
    const insW = SLIDE_W - 3.6;
    const insH = 0.62;
    s4.addShape(RECT, { x: insX, y: insY, w: insW, h: insH,
      fill: { color: NAVY }, line: { color: NAVY } });
    s4.addShape(RECT, { x: insX, y: insY, w: 0.06, h: insH,
      fill: { color: PINK }, line: { color: PINK } });
    s4.addText(
      'Consumers often respond on the 2nd or 3rd touch – not the first. Campaigns with consistent, well-timed follow-up typically see significantly higher overall engagement than single-touch outreach alone.',
      {
        x: insX + 0.18, y: insY + 0.05,
        w: insW - 0.26, h: insH - 0.1,
        fontSize: 9.5, color: WHITE, italic: true,
        fontFace: 'Aktiv Grotesk VF Medium', valign: 'middle'
      }
    );
  }

  if (cadence.cadenceOverallMedian != null) {
    const medStr = cadence.cadenceOverallMedian < 1
      ? `${(cadence.cadenceOverallMedian * 24).toFixed(1)} hours`
      : `${cadence.cadenceOverallMedian.toFixed(1)} days`;
    s4.addText(`Overall Median Re-Attempt Interval: ${medStr}  ·  Ideal range: 3–10 days`, {
      x: 0.4, y: SLIDE_H - 0.78,
      w: SLIDE_W - 0.8, h: 0.38,
      fontSize: 10, color: NAVY, bold: true, align: 'center',
      fontFace: 'Aktiv Grotesk VF Medium'
    });
  }
  s4.addNotes([
    'DELIVERY RE-ATTEMPT CADENCE — Speaker notes',
    '',
    'HOW TO READ THIS SLIDE:',
    'Shows the distribution of time gaps between consecutive delivery attempts on the same number. The "ideal" window (3–10 days) balances consumer recall (short enough to stay relevant) with respecting their decision time (long enough to avoid appearing aggressive).',
    '',
    'KEY TALKING POINTS:',
    '• Same-day re-attempts typically signal automated retry systems with no cadence control. These rarely convert and can contribute to carrier friction.',
    '• 1–2 day gaps are too soon for most consumers to have processed the first message. Treat these similarly to same-day.',
    '• 3–10 days is the proven sweet spot. Leads in this window convert at the highest incremental rate.',
    '• 11–15 days is acceptable but starts to lose the context of the first touch. Worth monitoring.',
    '• 16+ day gaps risk the consumer having forgotten about the first message entirely. Consider treating these as cold re-engagements.',
    '',
    'SINGLE-TOUCH OPPORTUNITY:',
    'Numbers that received only one attempt have never been retried. Even a modest re-attempt rate on this pool — at the right interval — can produce meaningful incremental deliveries from the same list investment.',
    '',
    'EXCEL REFERENCE: Retry Timing Analysis tab shows how gap length correlates with next-attempt success rate for this specific dataset.'
  ].join('\n'));
  slideFooter(s4);
  } // end includeSlideReAttemptCadence

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 5 – Recommended Actions (optional)
  // Cards with navy left accent bar (matches VoApps brand purple/pink motif)
  // ────────────────────────────────────────────────────────────────────────────
  if (includeSlideOpportunities) {
  const s5 = pptx.addSlide();
  s5.background = { color: CREAM };
  const s5HdrH = headerBar(pptx, s5, 'Opportunities to Maximize Performance', headerLogo, dateRangeStr);

  const noIssues = actions.length === 0
    || (actions.length === 1 && actions[0].includes('performance looks strong'));

  if (noIssues) {
    s5.addText('Campaign performance is strong across all key dimensions.\nDelivery rates, list health, and rotation patterns are all within recommended ranges.', {
      x: 0.5, y: 2.6, w: SLIDE_W - 1, h: 1.6,
      fontSize: 20, color: NAVY, bold: true, align: 'center', valign: 'middle',
      fontFace: 'IvyPresto Text'
    });
  } else {
    // ── Best Next Action banner ─────────────────────────────────────────────
    const bnaH = 0.98;
    const bnaY = s5HdrH + 0.30;  // 1.30 – tighter to header, more room for action rows
    s5.addShape(RECT, { x: 0.3, y: bnaY, w: SLIDE_W - 0.6, h: bnaH,
      fill: { color: NAVY }, line: { color: NAVY } });
    s5.addShape(RECT, { x: 0.3, y: bnaY, w: 0.08, h: bnaH,
      fill: { color: PINK }, line: { color: PINK } });
    s5.addText('BEST NEXT ACTION', {
      x: 0.5, y: bnaY + 0.17, w: SLIDE_W - 0.8, h: 0.22,
      fontSize: 10, bold: true, color: PINK_LIGHT,
      fontFace: 'Aktiv Grotesk VF Medium', charSpacing: 1
    });
    s5.addText(bestNextAction || '', {
      x: 0.5, y: bnaY + 0.38, w: SLIDE_W - 0.8, h: bnaH - 0.42,
      fontSize: 10, color: WHITE, italic: false,
      fontFace: 'Aktiv Grotesk VF Medium', valign: 'top'
    });

    const maxItems = Math.min(actions.length, 5); // one fewer since BNA takes space
    const availH   = SLIDE_H - s5HdrH - bnaH - 0.9;
    // Fill available height; cap only prevents overflow when 5 items are present
    const itemH    = Math.min(availH / Math.max(maxItems, 1), 1.40);
    let ay = bnaY + bnaH + 0.12;

    const slideNotes = [];
    for (let i = 0; i < maxItems; i++) {
      const raw      = actions[i];
      const colonIdx = raw.indexOf(':');
      const prefix   = colonIdx > -1 ? raw.substring(0, colonIdx + 1) : '';
      const rawBody  = colonIdx > -1 ? raw.substring(colonIdx + 1).trim() : raw;

      // Strip Excel tab cross-references from visible card text
      const { visible: body, notes: tabNotes } = separateTabRefs(rawBody);
      if (tabNotes.length) slideNotes.push(...tabNotes.map(n => `${prefix ? prefix + ' ' : ''}${n}`));

      const cardH = itemH - 0.08;

      // Card background – white with pink border
      s5.addShape(RECT, {
        x: 0.3, y: ay, w: SLIDE_W - 0.6, h: cardH,
        fill: { color: WHITE },
        line: { color: PINK_PALE, pt: 0.75 }
      });

      // Left accent bar – navy (matches brand motif from master deck)
      s5.addShape(RECT, {
        x: 0.3, y: ay, w: 0.08, h: cardH,
        fill: { color: PINK }, line: { color: PINK }
      });

      // Category label (bold navy) + body text
      const parts = prefix
        ? [
            { text: prefix + '  ', options: { bold: true, color: NAVY, fontFace: 'Aktiv Grotesk VF Medium' } },
            { text: body,           options: { color: TEXT_MID,  fontFace: 'Aktiv Grotesk VF Medium' } }
          ]
        : [{ text: body, options: { color: TEXT_MID, fontFace: 'Aktiv Grotesk VF Medium' } }];

      s5.addText(parts, {
        x: 0.52, y: ay + 0.06,
        w: SLIDE_W - 1.02, h: cardH - 0.15,
        fontSize: 9.5, valign: 'middle', wrap: true,
        fontFace: 'Aktiv Grotesk VF Medium'
      });

      ay += itemH;
    }

    if (actions.length > maxItems) {
      slideNotes.push(`+ ${actions.length - maxItems} additional recommendation(s) – see the Recommended Actions section in the Executive Summary tab of the Delivery Intelligence Excel report.`);
    }
    const s5Preamble = [
      'OPPORTUNITIES TO MAXIMIZE PERFORMANCE — Speaker notes',
      '',
      'PURPOSE OF THIS SLIDE:',
      'Surfaces the highest-leverage actions identified from the data. Each card is generated automatically from campaign metrics and is specific to this dataset — not generic advice.',
      '',
      'HOW TO PRESENT:',
      '• Walk through each recommendation in priority order (the most impactful action appears first).',
      '• Quantify where possible — e.g. "X numbers are single-touch; even a 20% re-attempt delivery rate adds Y successful deliveries."',
      '• Frame each recommendation as a business outcome, not a technical observation.',
      '',
      'EXCEL REPORT REFERENCES (for deeper analysis):'
    ].join('\n');
    if (slideNotes.length) {
      s5.addNotes(s5Preamble + '\n' + slideNotes.map((n, i) => `${i + 1}. ${n}`).join('\n'));
    } else {
      s5.addNotes(s5Preamble);
    }
  }
  slideFooter(s5);
  } // end includeSlideOpportunities

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 6 – Multi-Touch Delivery Funnel (when re-attempt data is available)
  // ────────────────────────────────────────────────────────────────────────────
  const rf = reAttemptData && reAttemptData.retentionFunnel;
  if (rf && rf.length >= 2) {
    const sfun = pptx.addSlide();
    sfun.background = { color: CREAM };
    const sfunHdrH = headerBar(pptx, sfun, 'Multi-Touch Delivery Funnel', headerLogo, dateRangeStr);

    const maxCount = rf[0].count;
    const maxBarW  = 9.0;
    const barH     = 0.72;
    const barGap   = 0.22;
    const levels   = Math.min(rf.length, 5);
    const totalH   = levels * barH + (levels - 1) * barGap;
    const startY   = sfunHdrH + (SLIDE_H - sfunHdrH - 0.35 - totalH) / 2;
    const labelW   = 2.0;
    const barX     = 0.3 + labelW + 0.10;

    // Sub-headline
    const multiPct = maxCount > 0 ? ((rf[1] ? rf[1].count : 0) / maxCount * 100).toFixed(0) : 0;
    sfun.addText(
      `${rf[1] ? rf[1].count.toLocaleString() : 0} of ${maxCount.toLocaleString()} numbers (${multiPct}%) received more than one delivery attempt`,
      {
        x: 0.3, y: sfunHdrH + 0.14, w: SLIDE_W - 0.6, h: 0.30,
        fontSize: 11, color: TEXT_SOFT, italic: true, align: 'center',
        fontFace: 'Aktiv Grotesk VF Medium'
      }
    );

    for (let i = 0; i < levels; i++) {
      const level = rf[i];
      const barW  = maxBarW * (level.count / maxCount);
      const bx    = barX; // left-aligned
      const by    = startY + i * (barH + barGap);
      const delivW = barW > 0 ? barW * (level.delivered / level.count) : 0;

      // Alternating purple shades for funnel depth effect
      const barColor  = i === 0 ? PURPLE       : i === 1 ? PURPLE_LIGHT
                      : i === 2 ? '9B94D4'     : i === 3 ? 'B2ACE2' : 'C8C3EC';
      const delivColor = i === 0 ? '0D053F'    : i === 1 ? PURPLE       : i === 2 ? PURPLE_LIGHT : '9B94D4';

      // Background track (full-width ghost bar for context)
      sfun.addShape(RECT, {
        x: barX, y: by, w: maxBarW, h: barH,
        fill: { color: 'EDE9F7' }, line: { color: 'DDD9EF', pt: 0.5 }
      });

      // Main bar (count at this attempt level)
      sfun.addShape(RECT, {
        x: bx, y: by, w: Math.max(0.04, barW), h: barH,
        fill: { color: barColor }, line: { color: barColor }
      });

      // Delivered portion overlay (darker shade, full bar height)
      if (delivW > 0.04) {
        sfun.addShape(RECT, {
          x: bx, y: by, w: delivW, h: barH,
          fill: { color: delivColor }, line: { color: delivColor }
        });
      }

      // Left label: "1+ Attempts" etc.
      sfun.addText(i === 0 ? '1st Attempt' : `${level.n}+ Attempts`, {
        x: 0.3, y: by, w: labelW, h: barH,
        fontSize: 12, bold: true, color: NAVY,
        fontFace: 'Aktiv Grotesk VF Medium', valign: 'middle', align: 'right'
      });

      // Count + delivered rate inside/beside bar
      const pctDel = level.count > 0 ? (level.delivered / level.count * 100).toFixed(0) : 0;
      // Use dark text on light bars (i>=2) so the label is always readable
      const labelColor = i < 2 ? WHITE : NAVY;
      sfun.addText(
        `${level.count.toLocaleString()}  ·  ${pctDel}% eventually delivered`,
        {
          x: bx + 0.14, y: by, w: Math.max(barW - 0.28, 2.0), h: barH,
          fontSize: 10.5, bold: false, color: labelColor,
          fontFace: 'Aktiv Grotesk VF Medium', valign: 'middle', align: 'left'
        }
      );
    }

    // Legend
    const legY = startY + levels * (barH + barGap) + 0.08;
    sfun.addShape(RECT, { x: barX, y: legY, w: 0.22, h: 0.16, fill: { color: PURPLE }, line: { color: PURPLE } });
    sfun.addText('Count at this attempt level', { x: barX + 0.28, y: legY - 0.01, w: 2.8, h: 0.18,
      fontSize: 9, color: TEXT_SOFT, fontFace: 'Aktiv Grotesk VF Medium' });
    sfun.addShape(RECT, { x: barX + 3.2, y: legY, w: 0.22, h: 0.16, fill: { color: NAVY }, line: { color: NAVY } });
    sfun.addText('Eventually delivered (result code 200)', { x: barX + 3.48, y: legY - 0.01, w: 3.4, h: 0.18,
      fontSize: 9, color: TEXT_SOFT, fontFace: 'Aktiv Grotesk VF Medium' });

    sfun.addNotes([
      'MULTI-TOUCH DELIVERY FUNNEL — Speaker notes',
      '',
      'HOW TO READ THIS CHART:',
      'Each bar represents the pool of numbers that reached at least that many attempts. The taller/wider the bar, the more numbers reached that attempt level.',
      '"Eventually delivered" means at least one successful DirectDrop Voicemail drop (result code 200) at any point — not necessarily on that specific attempt number.',
      'The darker segment of each bar shows how many of those numbers eventually delivered.',
      '',
      'KEY TALKING POINTS:',
      '• The gap between attempt 1 delivered % and later attempts quantifies the incremental value of re-attempts.',
      '• Diminishing returns typically set in around attempt 4–6. Numbers that have not delivered by then are candidates for suppression.',
      '• A large jump in eventual delivery rate between attempt 1 and 2 is a clear, quantifiable argument for multi-touch campaigns.',
      '',
      'EXCEL REFERENCE:',
      'The Attempt Funnel by Code tab in the Delivery Intelligence Excel report breaks this down by initial result code — showing which non-delivery outcomes are most worth retrying.'
    ].join('\n'));
    slideFooter(sfun);
  }

  // ────────────────────────────────────────────────────────────────────────────
  // CLOSING SLIDE – Let's Talk
  // Provocative discussion questions to drive client engagement
  // ────────────────────────────────────────────────────────────────────────────
  const sDx = pptx.addSlide();
  sDx.background = { color: NAVY };

  // Pink accent bar across the top
  sDx.addShape(RECT, {
    x: 0, y: 0, w: SLIDE_W, h: 0.08,
    fill: { color: PINK }, line: { color: PINK }
  });
  // Pink accent bar across the bottom
  sDx.addShape(RECT, {
    x: 0, y: SLIDE_H - 0.08, w: SLIDE_W, h: 0.08,
    fill: { color: PINK }, line: { color: PINK }
  });

  // "Let's Talk" headline
  sDx.addText("Let's Talk", {
    x: 0.5, y: 0.22, w: SLIDE_W - 1.0, h: 0.78,
    fontSize: 38, bold: true, color: WHITE,
    fontFace: 'IvyPresto Text', align: 'center', valign: 'middle'
  });

  // Thin pink divider below headline
  sDx.addShape(RECT, {
    x: SLIDE_W / 2 - 1.5, y: 0.95, w: 3.0, h: 0.04,
    fill: { color: PINK }, line: { color: PINK }
  });

  // Build data-driven, opportunity-focused discussion questions
  const cadenceSingleTouch = cadence?.cadenceSingleTouch || 0;

  const discussionQs = [
    cadenceSingleTouch > 0
      ? `${cadenceSingleTouch.toLocaleString()} numbers in this campaign were only touched once. Research consistently shows consumers need 2–3 touches before taking action — what would a well-timed follow-up campaign on this same list unlock for you?`
      : `What would a 10% lift in delivery rate mean for your program in concrete terms — more callbacks, more applications, more revenue? Let's map out what's realistic.`,

    staleWarmCount > 0
      ? `${staleWarmCount.toLocaleString()} consumers you've already successfully reached haven't been contacted in 30+ days — a pre-warmed audience with no cold-start problem. What's the highest-value message to put in front of them right now?`
      : `DirectDrop Voicemail delivered ${totalSuccess.toLocaleString()} messages this period — every one landing directly in a live voicemail inbox. What other parts of your consumer journey could benefit from that same frictionless, non-intrusive reach?`,

    agentHoursSaved > 0
      ? `This campaign freed an estimated ${agentHoursSaved.toLocaleString()} hours of agent capacity. If you could point that time at one specific business outcome, what would move the needle most for your team right now?`
      : `As DirectDrop Voicemail scales, the agent capacity it unlocks scales with it. What would your team do with an extra 100 hours a month freed from manual outreach?`,

    impliedRemovedCount > 0
      ? `~${impliedRemovedCount.toLocaleString()} numbers appear to have been removed after their last successful delivery — a strong signal of inbound callbacks already in motion. How are you capturing and converting those inbound calls when they come in?`
      : `DirectDrop Voicemail sits at the beginning of a consumer conversation. What happens on your end after the voicemail lands — and is there an opportunity to tighten that handoff?`,

    `If you could expand DirectDrop Voicemail into one new use case or channel integration — SMS follow-up, email sequencing, CRM automation — what would that look like, and what's one thing standing in the way today?`,
  ];

  const qStartY  = 1.10;
  const qH       = (SLIDE_H - qStartY - 0.28) / discussionQs.length;

  discussionQs.forEach((q, i) => {
    const qy = qStartY + i * qH;

    // Pink bullet diamond
    sDx.addShape(RECT, {
      x: 0.42, y: qy + qH / 2 - 0.08, w: 0.12, h: 0.12,
      fill: { color: PINK }, line: { color: PINK },
      rotate: 45
    });

    // Question text
    sDx.addText(q, {
      x: 0.70, y: qy, w: SLIDE_W - 1.0, h: qH,
      fontSize: 11.5, color: PINK_PALE,
      fontFace: 'Aktiv Grotesk VF Medium',
      valign: 'middle', wrap: true
    });
  });

  // ────────────────────────────────────────────────────────────────────────────
  sDx.addNotes([
    "LET'S TALK — Speaker notes",
    '',
    'PURPOSE OF THIS SLIDE:',
    'Close the review by opening a forward-looking conversation. Shift from "here\'s what the data shows" to "here\'s what we can do next together." Each question is anchored to a specific data point from this campaign.',
    '',
    'FACILITATION TIPS:',
    '• Lead with the question tied to the biggest opportunity surfaced in the data — typically the single-touch gap or the agent hours figure.',
    '• Frame every question as an opportunity, not a gap. The data is a starting point for expansion, not a report card.',
    '• Let the client answer fully before moving on — the best conversations come from listening, not presenting.',
    '• The implied callback figure (' + impliedRemovedCount.toLocaleString() + ' numbers) is often a surprise — clients don\'t always connect DirectDrop Voicemail delivery to inbound call volume.',
    '',
    'EXPANSION IDEAS TO SEED:',
    '• Multi-touch follow-up campaign on single-touch numbers',
    '• Re-engagement campaign for the stale warm audience',
    '• Integration with inbound call routing to capture callback intent',
    '• Extending DirectDrop Voicemail to new use cases: collections, appointment reminders, win-back, product launches',
    '• Combining with SMS or email for a coordinated multi-channel sequence'
  ].join('\n'));

  await pptx.writeFile({ fileName: outputPath });
}

module.exports = { generateBusinessReviewSlides };
