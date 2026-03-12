'use strict';

// businessReview.js — VoApps Delivery Intelligence business review slide generator
// Generates a branded .pptx alongside the Excel Delivery Intelligence Report

const pptxgen = require('pptxgenjs');
const path    = require('path');

// ─── Brand palette (from 2026 VoApps master deck theme) ───────────────────────
const NAVY         = '0D053F';   // dk2: deep navy — primary dark
const PINK         = 'FF4B7D';   // accent1: hot pink — primary accent
const PURPLE       = '3F2FB8';   // accent2: purple — secondary accent
const BLUE         = '16509B';   // accent3: blue — tertiary
const CHARCOAL     = '2E2C3E';   // accent4: dark charcoal
const PINK_PALE    = 'FAD6D7';   // accent5: pale pink
const PINK_LIGHT   = 'FF93B1';   // accent6: light pink
const PINK_MED     = 'FF6F97';   // mid pink — card accent variant
const PURPLE_LIGHT = '6558C6';   // lighter purple — card accent variant
const BLUE_LIGHT   = '4473AF';   // lighter blue — card accent variant
const CREAM        = 'FBF7F3';   // lt2: cream — body background
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
 * Matches the VoApps brand — navy background, white title text, pink icon block.
 * @param {string} [subheading] - Optional date range shown in PINK_LIGHT below the title.
 * @returns {number} Effective header height (use as base for content Y positioning).
 */
function headerBar(pptx, slide, title, logoPath, subheading) {
  const subH   = subheading ? 0.22 : 0;
  const totalH = HEADER_H + subH;

  // Full-width navy header (taller when subheading present)
  slide.addShape(RECT, {
    x: 0, y: 0, w: SLIDE_W, h: totalH,
    fill: { color: NAVY }, line: { color: NAVY }
  });

  // Pink icon square on the left — always original HEADER_H square, never stretched
  const iconBlockW = HEADER_H;
  slide.addShape(RECT, {
    x: 0, y: 0, w: iconBlockW, h: HEADER_H,
    fill: { color: PINK }, line: { color: PINK }
  });

  if (logoPath) {
    // Logo sits inside the pink block with a small inset
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

  if (subheading) {
    // Title sits in the original header band; subheading tucked just below it
    slide.addText(title, {
      x: titleX, y: 0, w: titleW, h: HEADER_H,
      fontSize: 22, bold: true, color: WHITE,
      fontFace: 'Aktiv Grotesk VF Medium',
      valign: 'middle', align: 'left',
      charSpacing: 0.5
    });
    slide.addText(subheading, {
      x: titleX, y: HEADER_H, w: titleW, h: subH,
      fontSize: 11, bold: false, color: PINK_LIGHT,
      fontFace: 'Aktiv Grotesk VF Medium',
      valign: 'middle', align: 'left'
    });
  } else {
    slide.addText(title, {
      x: titleX, y: 0, w: titleW, h: HEADER_H,
      fontSize: 22, bold: true, color: WHITE,
      fontFace: 'Aktiv Grotesk VF Medium',
      valign: 'middle', align: 'left',
      charSpacing: 0.5
    });
  }

  // Thin pink bottom accent line on the header
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
 * Metric card — cream background, colored top strip, label, large value, subtext.
 * Styled to match VoApps brand card pattern.
 */
function metricBox(slide, x, y, w, h, label, value, subtext, accentColor, valueFontSize) {
  const accent   = accentColor  || PINK;
  const valSize  = valueFontSize || 28;
  const stripH   = 0.06;
  // Value box height: shrinks to leave room for subtext when card is short
  const valBoxH  = subtext ? Math.min(0.68, h - 0.56) : 0.68;

  // Card background — cream
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

  // Label (small caps style — uppercase, spaced)
  slide.addText(label, {
    x: x + 0.16, y: y + stripH + 0.1,
    w: w - 0.32, h: 0.3,
    fontSize: 7.5, color: TEXT_SOFT,
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
      fontSize: 9, color: TEXT_SOFT,
      italic: false, align: 'left',
      fontFace: 'Aktiv Grotesk VF Medium'
    });
  }
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
async function generateBusinessReviewSlides(stats, outputPath, logoPath, squareLogo, circleLogo) {
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
    avgVariability,
    decayCurve,
    cadence,
    actions,
    bestNextAction,
    agentHoursSaved,
    staleWarmCount,
    minDate,
    maxDate,
    accountIds
  } = stats;

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
  // SLIDE 1 — Title
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

  // Lighter charcoal lower band for contrast
  s1.addShape(RECT, {
    x: 0, y: SLIDE_H - 1.4, w: SLIDE_W, h: 1.4,
    fill: { color: CHARCOAL }, line: { color: CHARCOAL }
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
    fontSize: 10, color: PINK_PALE,
    fontFace: 'Aktiv Grotesk VF Medium',
    align: 'center', italic: true
  });

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 2 — Campaign Overview
  // 2-row × 3-col metric cards on cream background
  // ────────────────────────────────────────────────────────────────────────────
  const s2 = pptx.addSlide();
  s2.background = { color: CREAM };
  const s2HdrH = headerBar(pptx, s2, 'Campaign Overview', headerLogo, dateRangeStr);
  const s2CY   = s2HdrH + 0.18;

  const bW   = 4.0;
  const bH   = 1.62;
  const bGap = 0.16;
  const row1Y = s2CY + 0.1;
  const row2Y = row1Y + bH + bGap;
  const c1 = 0.32;
  const c2 = c1 + bW + bGap;
  const c3 = c2 + bW + bGap;

  const successAccent = overallSuccessRate >= 75 ? BLUE
    : overallSuccessRate >= 50 ? PURPLE_LIGHT
    : CHARCOAL;
  const healthyPct = uniqueNumbers > 0 ? ((healthyCount / uniqueNumbers) * 100) : 0;
  const healthyAccent = healthyPct >= 80 ? BLUE : healthyPct >= 60 ? BLUE_LIGHT : PURPLE_LIGHT;

  metricBox(s2, c1, row1Y, bW, bH,
    'UNIQUE PHONE NUMBERS',
    uniqueNumbers.toLocaleString(),
    'Phone numbers analyzed in this report',
    PINK);

  metricBox(s2, c2, row1Y, bW, bH,
    'TOTAL DDVM ATTEMPTS',
    totalAttempts.toLocaleString(),
    'Delivery attempts recorded across all campaigns',
    NAVY);

  metricBox(s2, c3, row1Y, bW, bH,
    'OVERALL SUCCESS RATE',
    `${overallSuccessRate.toFixed(1)}%`,
    `${totalSuccess.toLocaleString()} voicemails successfully delivered`,
    successAccent);

  metricBox(s2, c1, row2Y, bW, bH,
    'SUCCESSFUL DELIVERIES',
    totalSuccess.toLocaleString(),
    `${overallSuccessRate.toFixed(1)}% of all attempts — voicemails delivered to consumers`,
    BLUE);

  metricBox(s2, c2, row2Y, bW, bH,
    'NUMBERS CONNECTING WELL',
    healthyCount.toLocaleString(),
    `${healthyPct.toFixed(0)}% of the list maintaining good delivery performance`,
    healthyAccent);

  // Date span — big day count, date range as subtext
  metricBox(s2, c3, row2Y, bW, bH,
    'DATE SPAN',
    daySpan > 0 ? `${daySpan} Days` : dateRangeStr,
    daySpan > 0 ? dateRangeStr : '',
    PINK_LIGHT, 32);

  // ── Single-touch opportunity callout strip ────────────────────────────────
  const cadenceTotalNumbers = (cadence.cadenceSingleTouch || 0) + (cadence.cadenceMultiTouchCount || 0);
  if (cadenceTotalNumbers > 0) {
    const stPct = (cadence.cadenceSingleTouch / cadenceTotalNumbers * 100).toFixed(1);
    const stripY = row2Y + bH + 0.20;
    const stripH = 0.68;
    const stripX = c1;
    const stripW = SLIDE_W - c1 * 2;

    // Cream card with pink left accent bar
    s2.addShape(RECT, { x: stripX, y: stripY, w: stripW, h: stripH,
      fill: { color: WHITE }, line: { color: PINK_PALE, pt: 1 } });
    s2.addShape(RECT, { x: stripX, y: stripY, w: 0.07, h: stripH,
      fill: { color: PINK }, line: { color: PINK } });

    // Bold stat left
    s2.addText(
      `${cadence.cadenceSingleTouch.toLocaleString()} numbers (${stPct}%) received only one DDVM attempt this period.`,
      {
        x: stripX + 0.18, y: stripY + 0.04,
        w: stripW - 0.26, h: 0.28,
        fontSize: 11, bold: true, color: NAVY,
        fontFace: 'Aktiv Grotesk VF Medium', valign: 'middle'
      }
    );
    // Supporting text — blend single-touch + stale warm into one insight line
    const staleNote = (staleWarmCount > 0)
      ? ` Additionally, ${staleWarmCount.toLocaleString()} numbers with prior successful deliveries haven't been contacted in 30+ days — confirmed-reachable low-hanging fruit for re-engagement.`
      : '';
    s2.addText(
      `Consumers often need 2–3 touches before taking action. A follow-up campaign at a 3–10 day interval can produce meaningful incremental results from this same list.${staleNote}`,
      {
        x: stripX + 0.18, y: stripY + 0.33,
        w: stripW - 0.26, h: 0.28,
        fontSize: 9.5, color: TEXT_SOFT, italic: true,
        fontFace: 'Aktiv Grotesk VF Medium', valign: 'middle'
      }
    );
  }

  // ── Agent Hours Saved — full-width card below callout strip ─────────────────
  if (agentHoursSaved > 0) {
    const ahCardH = 1.0;
    // ahY derived from strip bottom so it's always accurate regardless of header height
    const ahStripBottom = row2Y + bH + 0.20 + 0.68; // stripY + stripH
    const ahY = cadenceTotalNumbers > 0
      ? ahStripBottom + 0.10
      : row2Y + bH + 0.22;
    metricBox(s2, c1, ahY, SLIDE_W - c1 * 2, ahCardH,
      'AGENT HOURS SAVED (EST.)',
      `${agentHoursSaved.toLocaleString()} hrs`,
      `Based on ${totalSuccess.toLocaleString()} successful deliveries × 3 min avg manual voicemail handle time — capacity your agents didn't need to spend on outreach`,
      PURPLE, 28);
  }

  slideFooter(s2);

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 3 — Success Probability by Attempt
  // ────────────────────────────────────────────────────────────────────────────
  const s3 = pptx.addSlide();
  s3.background = { color: CREAM };
  const s3HdrH = headerBar(pptx, s3, 'Success Probability by Attempt', headerLogo, dateRangeStr);
  const s3CY   = s3HdrH + 0.18;

  s3.addText(
    'Each row represents all numbers at that attempt count. As attempt index rises, the pool shifts toward harder-to-reach numbers — success probability naturally declines.',
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
      const insight  = pct >= 50 ? 'Good — Continue'
        : pct >= 25 ? 'Declining — Monitor'
        : pct >= 15 ? 'Low — Review'
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

  s3.addText(
    'Numbers with 4–6+ consecutive failures and low success rates are listed in the Suppression Candidates tab of the report.',
    {
      x: 0.4, y: SLIDE_H - 0.78,
      w: SLIDE_W - 0.8, h: 0.36,
      fontSize: 8.5, color: TEXT_SOFT, italic: true, align: 'center',
      fontFace: 'Aktiv Grotesk VF Medium'
    }
  );
  slideFooter(s3);

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 4 — Delivery Re-Attempt Cadence
  // ────────────────────────────────────────────────────────────────────────────
  const s4 = pptx.addSlide();
  s4.background = { color: CREAM };
  const s4HdrH = headerBar(pptx, s4, 'Delivery Re-Attempt Cadence', headerLogo, dateRangeStr);
  const s4CY   = s4HdrH + 0.18;

  const totalMulti = cadence.cadenceMultiTouchCount || 1;
  const s4TotalNumbers = (cadence.cadenceSingleTouch || 0) + (cadence.cadenceMultiTouchCount || 0);
  const s4SinglePct = s4TotalNumbers > 0 ? (cadence.cadenceSingleTouch / s4TotalNumbers * 100).toFixed(1) : '0.0';
  const s4MultiPct  = s4TotalNumbers > 0 ? (cadence.cadenceMultiTouchCount / s4TotalNumbers * 100).toFixed(1) : '0.0';
  s4.addText(
    `${cadence.cadenceMultiTouchCount.toLocaleString()} numbers (${s4MultiPct}%) had 2+ delivery attempts. ${cadence.cadenceSingleTouch.toLocaleString()} numbers (${s4SinglePct}%) were contacted only once — each a potential opportunity for an additional touch. Breakdown by median interval between consecutive attempts:`,
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
      'Consumers often respond on the 2nd or 3rd touch — not the first. Campaigns with consistent, well-timed follow-up typically see significantly higher overall engagement than single-touch outreach alone.',
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

  // ────────────────────────────────────────────────────────────────────────────
  // SLIDE 5 — Recommended Actions
  // Cards with navy left accent bar (matches VoApps brand purple/pink motif)
  // ────────────────────────────────────────────────────────────────────────────
  const s5 = pptx.addSlide();
  s5.background = { color: CREAM };
  const s5HdrH = headerBar(pptx, s5, 'Opportunities to Maximize Performance', headerLogo, dateRangeStr);
  const s5CY   = s5HdrH + 0.18;

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
    const bnaH = 0.78;
    const bnaY = s5CY + 0.1;
    s5.addShape(RECT, { x: 0.3, y: bnaY, w: SLIDE_W - 0.6, h: bnaH,
      fill: { color: NAVY }, line: { color: NAVY } });
    s5.addShape(RECT, { x: 0.3, y: bnaY, w: 0.08, h: bnaH,
      fill: { color: PINK }, line: { color: PINK } });
    s5.addText('BEST NEXT ACTION', {
      x: 0.5, y: bnaY + 0.04, w: SLIDE_W - 0.8, h: 0.22,
      fontSize: 8.5, bold: true, color: PINK_LIGHT,
      fontFace: 'Aktiv Grotesk VF Medium', charSpacing: 1
    });
    s5.addText(bestNextAction || '', {
      x: 0.5, y: bnaY + 0.25, w: SLIDE_W - 0.8, h: bnaH - 0.3,
      fontSize: 10, color: WHITE, italic: false,
      fontFace: 'Aktiv Grotesk VF Medium', valign: 'top'
    });

    const maxItems = Math.min(actions.length, 5); // one fewer since BNA takes space
    const availH   = SLIDE_H - s5HdrH - bnaH - 1.2;
    const itemH    = Math.min(0.82, availH / Math.max(maxItems, 1));
    let ay = bnaY + bnaH + 0.12;

    for (let i = 0; i < maxItems; i++) {
      const raw      = actions[i];
      const colonIdx = raw.indexOf(':');
      const prefix   = colonIdx > -1 ? raw.substring(0, colonIdx + 1) : '';
      const body     = colonIdx > -1 ? raw.substring(colonIdx + 1).trim() : raw;

      const cardH = itemH - 0.08;

      // Card background — white with pink border
      s5.addShape(RECT, {
        x: 0.3, y: ay, w: SLIDE_W - 0.6, h: cardH,
        fill: { color: WHITE },
        line: { color: PINK_PALE, pt: 0.75 }
      });

      // Left accent bar — navy (matches brand motif from master deck)
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
      s5.addText(`+ ${actions.length - maxItems} more — see the Recommended Actions section in the Executive Summary tab.`, {
        x: 0.4, y: SLIDE_H - 0.6,
        w: SLIDE_W - 0.8, h: 0.35,
        fontSize: 9, color: TEXT_SOFT, italic: true, align: 'center',
        fontFace: 'Aktiv Grotesk VF Medium'
      });
    }
  }
  slideFooter(s5);

  // ────────────────────────────────────────────────────────────────────────────
  await pptx.writeFile({ fileName: outputPath });
}

module.exports = { generateBusinessReviewSlides };
