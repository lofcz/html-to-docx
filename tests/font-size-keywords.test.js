/**
 * CSS font-size keyword support + inheritance tests.
 *
 * Covers:
 *  - CSS absolute keywords (xx-small ... xxx-large) resolving to spec pixel
 *    values converted to Word half-points.
 *  - CSS relative keywords (`smaller`, `larger`) resolving against the NEAREST
 *    ANCESTOR's font-size, not the document default. This is the fix for the
 *    font-inheritance bug flagged on upstream issue/PR #15.
 *  - Percentage font-size also resolving against the inherited ancestor size.
 *
 * Word stores font size as half-points in `<w:sz w:val="N"/>` (e.g. 11pt = 22).
 */

import HTMLtoDOCX from '../index.js';
import { parseDOCX } from './helpers/docx-assertions.js';

// Pull <w:sz w:val="N"/> from the run that wraps `needle` text.
function extractRunFontSize(xml, needle) {
  const runRe = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;
  let m;
  while ((m = runRe.exec(xml)) !== null) {
    if (m[1].includes(needle)) {
      const sz = m[1].match(/<w:sz w:val="(\d+)"/);
      return sz ? parseInt(sz[1], 10) : null;
    }
  }
  return null;
}

describe('CSS font-size absolute keywords', () => {
  test.each([
    ['xx-small', 9],
    ['x-small', 10],
    ['small', 13],
    ['medium', 16],
    ['large', 18],
    ['x-large', 24],
    ['xx-large', 32],
    ['xxx-large', 48],
  ])('%s maps to the spec px value in half-points', async (keyword, px) => {
    const html = `<p><span style="font-size: ${keyword};">${keyword}-text</span></p>`;
    const buf = await HTMLtoDOCX(html, null, { deterministicIds: true });
    const { xml } = await parseDOCX(buf);
    // pixelToHIP for these whole-px values: px -> half-points via the existing
    // converter. We compute the expectation the same way the library does.
    const expected = Math.round(px * 72 * 2 / 96); // px -> pt (*72/96) -> HIP (*2)
    expect(extractRunFontSize(xml, `${keyword}-text`)).toBe(expected);
  });
});

describe('CSS relative font-size keywords inherit from parent (issue #15 fix)', () => {
  test('smaller resolves against the parent font-size, not the document default', async () => {
    // Parent is 16px. `smaller` = parent * 5/6.
    // Document default is 11pt (22 half-points); the bug was using that default.
    const html =
      '<p style="font-size: 16px;">' +
      '<span style="font-size: smaller;">smaller-child</span>' +
      '</p>';
    const buf = await HTMLtoDOCX(html, null, { deterministicIds: true });
    const { xml } = await parseDOCX(buf);

    const parentPx = 16;
    const parentHIP = Math.round((parentPx * 72 * 2) / 96); // 24 half-points
    const expected = Math.round(parentHIP * (5 / 6)); // 20 half-points
    expect(extractRunFontSize(xml, 'smaller-child')).toBe(expected);
    // And critically NOT the buggy document-default-derived value:
    // Math.round(22 * 5/6) === 18, so 20 !== 18 proves inheritance works.
    expect(extractRunFontSize(xml, 'smaller-child')).not.toBe(Math.round((22 * 5) / 6));
  });

  test('larger resolves against the parent font-size, not the document default', async () => {
    const html =
      '<p style="font-size: 16px;">' +
      '<span style="font-size: larger;">larger-child</span>' +
      '</p>';
    const buf = await HTMLtoDOCX(html, null, { deterministicIds: true });
    const { xml } = await parseDOCX(buf);

    const parentPx = 16;
    const parentHIP = Math.round((parentPx * 72 * 2) / 96); // 24
    const expected = Math.round(parentHIP * (6 / 5)); // ~29 half-points
    expect(extractRunFontSize(xml, 'larger-child')).toBe(expected);
  });

  test('nested relative keywords chain through ancestors', async () => {
    // 16px parent -> larger (~1.2x) -> smaller (~5/6 of that).
    const html =
      '<p style="font-size: 16px;">' +
      '<span style="font-size: larger;">' +
      '<span style="font-size: smaller;">nested-child</span>' +
      '</span>' +
      '</p>';
    const buf = await HTMLtoDOCX(html, null, { deterministicIds: true });
    const { xml } = await parseDOCX(buf);

    const parentHIP = Math.round((16 * 72 * 2) / 96); // 24
    const afterLarger = Math.round(parentHIP * (6 / 5)); // 29
    const afterSmaller = Math.round(afterLarger * (5 / 6)); // 24
    expect(extractRunFontSize(xml, 'nested-child')).toBe(afterSmaller);
  });
});

describe('Percentage font-size resolves against inherited ancestor', () => {
  test('50% of a 24px parent (not the document default)', async () => {
    const html =
      '<p style="font-size: 24px;">' +
      '<span style="font-size: 50%;">half-child</span>' +
      '</p>';
    const buf = await HTMLtoDOCX(html, null, { deterministicIds: true });
    const { xml } = await parseDOCX(buf);

    const parentHIP = Math.round((24 * 72 * 2) / 96); // 36
    const expected = Math.round((50 * parentHIP) / 100); // 18
    expect(extractRunFontSize(xml, 'half-child')).toBe(expected);
  });
});
