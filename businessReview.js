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
    reAttemptData                = null
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

  if (acctList) {
    s1.addText(`Account${accountIds.length > 1 ? 's' : ''}: ${acctList}`, {
      x: 0.5, y: 4.5, w: SLIDE_W - 1.0, h: 0.5,
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

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 2 – Delivery Performance Gauge
  // Half-donut speedometer + three key stats
  // ────────────────────────────────────────────────────────────────────────────
  const sGauge = pptx.addSlide();
  sGauge.background = { color: CREAM };
  const sGaugeHdrH = headerBar(pptx, sGauge, 'Delivery Performance Snapshot', headerLogo, dateRangeStr);

  // ── Three thermometers ───────────────────────────────────────────────────
  // Metric 1: Overall delivery rate (pink)
  // Metric 2: Healthy numbers % of total unique (purple/blue)
  // Metric 3: Clean list % = numbers NOT flagged as suppression candidates (amber→green)
  const healthyPct      = uniqueNumbers > 0 ? (healthyCount      / uniqueNumbers) * 100 : 0;
  const cleanPct        = uniqueNumbers > 0 ? ((uniqueNumbers - suppressionCandidateCount) / uniqueNumbers) * 100 : 100;

  const delivColor = overallSuccessRate >= 75 ? PINK   : overallSuccessRate >= 50 ? PURPLE : CHARCOAL;
  const healthColor = healthyPct >= 60 ? '2D8A60' : healthyPct >= 40 ? PURPLE : CHARCOAL;
  const cleanColor  = cleanPct  >= 75 ? '2D8A60' : cleanPct  >= 50 ? 'C97A20' : RED;

  const thTubeH = 3.1;
  const thTubeW = 0.52;
  const thTopY  = sGaugeHdrH + 0.35;
  const colCenters = [SLIDE_W * (1/6), SLIDE_W * (3/6), SLIDE_W * (5/6)];

  drawThermometer(sGauge, colCenters[0], thTopY, thTubeH, thTubeW,
    overallSuccessRate, delivColor,
    'DELIVERY RATE', `${overallSuccessRate.toFixed(1)}%`);

  drawThermometer(sGauge, colCenters[1], thTopY, thTubeH, thTubeW,
    healthyPct, healthColor,
    'NUMBERS WITH A DELIVERY', `${healthyPct.toFixed(1)}%`);

  drawThermometer(sGauge, colCenters[2], thTopY, thTubeH, thTubeW,
    cleanPct, cleanColor,
    'CLEAN LIST (NOT FLAGGED)', `${cleanPct.toFixed(1)}%`);

  // ── Stats row ────────────────────────────────────────────────────────────
  const gcX  = 0.5;
  const gcW  = SLIDE_W - 1.0;
  // Place stats row below the thermometer bulbs
  const bulbBottomY = thTopY + thTubeH + thTubeW * 0.9 * 2 + 0.18;
  const statY = Math.max(bulbBottomY + 0.55, SLIDE_H - 1.45);

  sGauge.addShape(RECT, {
    x: gcX, y: statY - 0.04, w: gcW, h: 0.02,
    fill: { color: 'E0DCEA' }, line: { color: 'E0DCEA' }
  });

  // Three key stats in a row below the thermometers
  const gaugeStats = [
    { value: totalAttempts.toLocaleString(),              label: 'TOTAL ATTEMPTS'  },
    { value: totalSuccess.toLocaleString(),               label: 'DELIVERED'       },
    { value: (totalAttempts - totalSuccess).toLocaleString(), label: 'NOT DELIVERED' }
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

  // ── Card registry – all possible metric cards ─────────────────────────────
  // firstAttemptSuccessRate uses decayCurve[0] if available
  const firstAttemptRate = (decayCurve && decayCurve.length > 0)
    ? `${(decayCurve[0].probability * 100).toFixed(1)}%`
    : `${overallSuccessRate.toFixed(1)}%`;

  const ALL_CARDS = {
    uniquePhoneNumbers:    { label: 'UNIQUE PHONE NUMBERS',      value: uniqueNumbers.toLocaleString(),                              sub: 'Phone numbers analyzed in this report',                               accent: PINK },
    totalAttempts:         { label: 'TOTAL DDVM ATTEMPTS',       value: totalAttempts.toLocaleString(),                             sub: 'Delivery attempts recorded across all campaigns',                     accent: NAVY },
    overallSuccessRate:    { label: 'OVERALL SUCCESS RATE',      value: `${overallSuccessRate.toFixed(1)}%`,                        sub: `${totalSuccess.toLocaleString()} voicemails successfully delivered`,  accent: successAccent },
    successfulDeliveries:  { label: 'SUCCESSFUL DELIVERIES',     value: totalSuccess.toLocaleString(),                              sub: `${overallSuccessRate.toFixed(1)}% of all attempts – voicemails delivered to consumers`, accent: BLUE },
    numbersConnectingWell: { label: 'NUMBERS CONNECTING WELL',   value: healthyCount.toLocaleString(),                              sub: `${healthyPct.toFixed(0)}% of the list maintaining good delivery performance`, accent: healthyAccent },
    dateSpan:              { label: 'DATE SPAN',                 value: daySpan > 0 ? `${daySpan} Days` : dateRangeStr,            sub: daySpan > 0 ? dateRangeStr : '',                                       accent: PINK_LIGHT, fontSize: 28 },
    agentHoursSaved:       { label: 'AGENT HOURS SAVED (EST.)',  value: agentHoursSaved > 0 ? `${agentHoursSaved.toLocaleString()} hrs` : '—', sub: `Based on ${totalSuccess.toLocaleString()} deliveries × 3 min avg handle time`, accent: PURPLE },
    unsuccessfulAttempts:  { label: 'UNSUCCESSFUL ATTEMPTS',     value: (totalAttempts - totalSuccess).toLocaleString(),           sub: `${(100 - overallSuccessRate).toFixed(1)}% of all attempts`,           accent: CHARCOAL },
    firstAttemptSuccessRate: { label: 'FIRST ATTEMPT SUCCESS RATE', value: firstAttemptRate,                                       sub: 'Success rate on the very first delivery attempt to each number',      accent: BLUE_LIGHT },
    avgAttemptsPerNumber:  { label: 'AVG ATTEMPTS PER NUMBER',   value: uniqueNumbers > 0 ? (totalAttempts / uniqueNumbers).toFixed(1) : '—', sub: 'Average total delivery attempts per unique phone number',    accent: PURPLE_LIGHT },
    nonDeliverableNumbers:   { label: 'LIKELY NON-DELIVERABLE',      value: (suppressionCandidateCount || 0).toLocaleString(),         sub: 'Numbers meeting suppression criteria – repeated failures over an extended span',         accent: CHARCOAL },
    impliedRemovedNumbers:   { label: 'IMPLIED REMOVED AFTER DELIVERY', value: impliedRemovedCount.toLocaleString(),                    sub: 'Delivered numbers not re-attempted after one cadence window – likely suppressed/removed from list', accent: PURPLE },
    impliedCallbackOppty:    { label: 'IMPLIED CALLBACK RATE',         value: `${impliedCallbackRate.toFixed(1)}%`,                     sub: `${impliedRemovedCount.toLocaleString()} delivered numbers appear removed – each is a potential inbound callback`, accent: BLUE_LIGHT }
  };

  const DEFAULT_CARDS = [
    'uniquePhoneNumbers', 'totalAttempts', 'overallSuccessRate',
    'successfulDeliveries', 'numbersConnectingWell', 'dateSpan'
  ];
  // Allow up to 12 cards across two overview slides (6 per slide)
  const cardKeys = (Array.isArray(overviewCards) && overviewCards.length > 0)
    ? overviewCards.slice(0, 12)
    : DEFAULT_CARDS;

  const cardPositions = [
    [c1, row1Y], [c2, row1Y], [c3, row1Y],
    [c1, row2Y], [c2, row2Y], [c3, row2Y]
  ];

  // Helper: render a page of up to 6 cards onto a slide
  function renderCardPage(slide, pageKeys) {
    pageKeys.forEach((key, i) => {
      const card = ALL_CARDS[key];
      const pos  = cardPositions[i];
      if (!card || !pos) return;
      if (key === 'overallSuccessRate') {
        const [cx, cy] = pos;
        slide.addShape(RECT, { x: cx, y: cy, w: bW, h: bH, fill: { color: CREAM }, line: { color: PINK_PALE, pt: 1 } });
        slide.addShape(RECT, { x: cx, y: cy, w: bW, h: 0.06, fill: { color: card.accent }, line: { color: card.accent } });
        slide.addText(card.label, { x: cx + 0.16, y: cy + 0.16, w: bW - 0.32, h: 0.26,
          fontSize: 11, color: TEXT_SOFT, bold: false, fontFace: 'Aktiv Grotesk VF Medium', charSpacing: 1.5 });
        slide.addText(card.value, { x: cx + 0.16, y: cy + 0.42, w: bW - 0.32, h: 0.56,
          fontSize: 28, bold: true, color: NAVY, fontFace: 'IvyPresto Text', align: 'left', valign: 'top', shrinkText: true });
        const pbX = cx + 0.16, pbY = cy + 1.02, pbW2 = bW - 0.32, pbH = 0.14;
        slide.addShape(RECT, { x: pbX, y: pbY, w: pbW2, h: pbH, fill: { color: 'DDD9EF' }, line: { color: 'DDD9EF' } });
        const fillW = Math.max(0.04, pbW2 * (overallSuccessRate / 100));
        slide.addShape(RECT, { x: pbX, y: pbY, w: fillW, h: pbH, fill: { color: card.accent }, line: { color: card.accent } });
        slide.addText(card.sub, { x: cx + 0.16, y: cy + 1.20, w: bW - 0.32, h: 0.30,
          fontSize: 10, color: TEXT_SOFT, fontFace: 'Aktiv Grotesk VF Medium' });
      } else {
        metricBox(slide, pos[0], pos[1], bW, bH, card.label, card.value, card.sub, card.accent, card.fontSize || 32);
      }
    });
  }

  // Page 1 (always exists)
  renderCardPage(s2, cardKeys.slice(0, 6));

  // Page 2 (only if there are more than 6 selected cards)
  let s2b = null;
  if (cardKeys.length > 6) {
    s2b = pptx.addSlide();
    s2b.background = { color: CREAM };
    headerBar(pptx, s2b, 'High-Level Overview (cont.)', headerLogo, dateRangeStr);
    renderCardPage(s2b, cardKeys.slice(6, 12));
  }

  // ── Single-touch opportunity callout strip (on last overview slide) ─────────
  const lastOverviewSlide = s2b || s2;
  const cadenceTotalNumbers = (cadence.cadenceSingleTouch || 0) + (cadence.cadenceMultiTouchCount || 0);
  if (cadenceTotalNumbers > 0) {
    const stPct = (cadence.cadenceSingleTouch / cadenceTotalNumbers * 100).toFixed(1);
    const stripY = row2Y + bH + 0.20;
    const stripH = 0.68;
    const stripX = c1;
    const stripW = 3 * bW + 2 * bGap;

    lastOverviewSlide.addShape(RECT, { x: stripX, y: stripY, w: stripW, h: stripH,
      fill: { color: WHITE }, line: { color: PINK_PALE, pt: 1 } });
    lastOverviewSlide.addShape(RECT, { x: stripX, y: stripY, w: 0.07, h: stripH,
      fill: { color: PINK }, line: { color: PINK } });

    lastOverviewSlide.addText(
      `${cadence.cadenceSingleTouch.toLocaleString()} numbers (${stPct}%) received only one DDVM attempt this period.`,
      { x: stripX + 0.18, y: stripY + 0.10, w: stripW - 0.26, h: 0.28,
        fontSize: 12, bold: true, color: NAVY, fontFace: 'Aktiv Grotesk VF Medium', valign: 'middle' }
    );
    const staleNote = staleWarmCount > 0
      ? ` Additionally, ${staleWarmCount.toLocaleString()} numbers with prior successful deliveries haven't been contacted in 30+ days – confirmed-reachable low-hanging fruit for re-engagement.`
      : '';
    const impliedNote = impliedRemovedCount > 0
      ? ` ~${impliedRemovedCount.toLocaleString()} numbers appear to have been removed after their last successful delivery – a potential inbound callback opportunity.`
      : '';
    lastOverviewSlide.addText(
      `Consumers often need 2\u20133 touches before taking action. A follow-up campaign at a 3\u201310 day interval can produce meaningful incremental results from this same list.${staleNote}${impliedNote}`,
      { x: stripX + 0.18, y: stripY + 0.29, w: stripW - 0.26, h: 0.28,
        fontSize: 10, color: TEXT_SOFT, italic: true, fontFace: 'Aktiv Grotesk VF Medium', valign: 'middle' }
    );
  }

  // ── Agent Hours Saved – full-width purple card below callout strip ──────────
  // Only render as banner if it's not already shown as one of the metric cards
  if (agentHoursSaved > 0 && !cardKeys.includes('agentHoursSaved')) {
    const ahCardH  = 1.14;
    const ahCardW  = 3 * bW + 2 * bGap;
    const ahStripBottom = row2Y + bH + 0.20 + 0.68;
    const ahY = cadenceTotalNumbers > 0
      ? ahStripBottom + 0.10
      : row2Y + bH + 0.22;

    const PURPLE_BG    = 'EBE9F7';
    const PURPLE_STRIP = '3F2FB8';

    lastOverviewSlide.addShape(RECT, { x: c1, y: ahY, w: ahCardW, h: ahCardH, fill: { color: PURPLE_BG }, line: { color: PURPLE_BG } });
    lastOverviewSlide.addShape(RECT, { x: c1, y: ahY, w: ahCardW, h: 0.05, fill: { color: PURPLE_STRIP }, line: { color: PURPLE_STRIP } });
    lastOverviewSlide.addText('AGENT HOURS SAVED (EST.)', {
      x: c1 + 0.16, y: ahY + 0.127, w: ahCardW - 0.32, h: 0.28,
      fontSize: 11, color: TEXT_SOFT, bold: false, fontFace: 'Aktiv Grotesk VF Medium', align: 'left', charSpacing: 1.5
    });
    lastOverviewSlide.addText(`${agentHoursSaved.toLocaleString()} hrs`, {
      x: c1 + 0.16, y: ahY + 0.308, w: ahCardW - 0.32, h: 0.44,
      fontSize: 32, bold: true, color: NAVY, fontFace: 'IvyPresto Text', align: 'left', valign: 'top', shrinkText: true
    });
    lastOverviewSlide.addText(
      `Based on ${totalSuccess.toLocaleString()} successful deliveries \u00d7 3 min avg manual voicemail handle time \u2014 capacity your agents didn\u2019t need to spend on outreach`,
      { x: c1 + 0.16, y: ahY + 0.784, w: ahCardW - 0.32, h: 0.3, fontSize: 10, color: TEXT_SOFT, fontFace: 'Aktiv Grotesk VF Medium', align: 'left' }
    );
  }

  slideFooter(s2);
  if (s2b) slideFooter(s2b);

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
      fontSize: 9.5, color: TEXT_SOFT, italic: true,
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

  s3.addTable(tblRows, {
    x: 1.4, y: s3CY + 0.58,
    w: SLIDE_W - 2.8,
    fontSize: 10.5,
    rowH: 0.43,
    border: { type: 'solid', color: PINK_PALE, pt: 0.75 }
  });

  s3.addNotes('For your reference: Numbers with 4–6+ consecutive failures and low success rates are listed in the Suppression Candidates tab of the Delivery Intelligence Excel report.');
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
    `${cadence.cadenceMultiTouchCount.toLocaleString()} numbers (${s4MultiPct}%) had 2+ delivery attempts – across all result types (successfully delivered, unsuccessful, voicemail not setup, voicemail full, and not in service). ${cadence.cadenceSingleTouch.toLocaleString()} numbers (${s4SinglePct}%) were contacted only once – each a potential opportunity for an additional touch. Breakdown by median interval between consecutive attempts:`,
    {
      x: 0.4, y: s4CY + 0.08,
      w: SLIDE_W - 0.8, h: 0.44,
      fontSize: 9.5, color: TEXT_SOFT, italic: true,
      fontFace: 'Aktiv Grotesk VF Medium'
    }
  );

  const cadenceRows = [
    { label: 'Same-day re-attempt  (< 1 day)', count: cadence.cadenceBucket_sameDay,  ideal: false, warn: true,  note: false, long: false },
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

  s4.addTable(cTbl, {
    x: 1.8, y: s4CY + 0.58,
    w: SLIDE_W - 3.6,
    fontSize: 10.5, rowH: 0.44,
    border: { type: 'solid', color: PINK_PALE, pt: 0.75 }
  });

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
    const availH   = SLIDE_H - s5HdrH - bnaH - 1.2;
    const itemH    = Math.min(0.82, availH / Math.max(maxItems, 1));
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
    if (slideNotes.length) {
      s5.addNotes('Speaker notes – Excel report references:\n' + slideNotes.map((n, i) => `${i + 1}. ${n}`).join('\n'));
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
      sfun.addText(
        `${level.count.toLocaleString()}  ·  ${pctDel}% eventually delivered`,
        {
          x: bx + 0.14, y: by, w: Math.max(barW - 0.28, 2.0), h: barH,
          fontSize: 10.5, bold: false, color: i === 0 ? WHITE : WHITE,
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
    sfun.addText('Eventually delivered (200)', { x: barX + 3.48, y: legY - 0.01, w: 2.8, h: 0.18,
      fontSize: 9, color: TEXT_SOFT, fontFace: 'Aktiv Grotesk VF Medium' });

    sfun.addNotes('Data source: Attempt Funnel by Code tab in the Delivery Intelligence Excel report.\n' +
      '• Bar width shows how many numbers reached that attempt level (narrowing = fewer numbers retried that many times).\n' +
      '• "Eventually delivered" = had at least one successful drop (200) at any point.');
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
      ? `Research shows consumers typically need 2–3 touches before taking action. Of the ${cadenceSingleTouch.toLocaleString()} numbers that received just one attempt this period, how many could convert with a well-timed follow-up?`
      : `What does your re-attempt strategy look like today – and is there an opportunity to capture more value from the same list with a structured multi-touch cadence?`,

    staleWarmCount > 0
      ? `${staleWarmCount.toLocaleString()} consumers you've already successfully reached haven't been contacted in 30+ days – a pre-warmed audience ready for re-engagement. What's the right message to put in front of them right now?`
      : `How quickly do you follow up after a successful delivery? A timely second touch to already-engaged consumers tends to see significantly higher conversion rates.`,

    `How recently has your active list been refreshed? Periodic list maintenance is one of the highest-leverage ways to raise delivery rates – what does that process look like for your team today?`,

    agentHoursSaved > 0
      ? `DDVM freed up an estimated ${agentHoursSaved.toLocaleString()} hours of agent time this period. How is that capacity being put to work – and where would redirecting it create the most impact for your business?`
      : `As DDVM scales, so does the agent capacity it unlocks. How are you thinking about reinvesting that time into higher-value conversations or expanded outreach?`,

    `If you could change one thing about how this program runs today, what would it be – and what's been stopping you from making that change?`,
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
  await pptx.writeFile({ fileName: outputPath });
}

module.exports = { generateBusinessReviewSlides };
