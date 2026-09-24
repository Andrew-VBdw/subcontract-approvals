/*
 * Dawson Wallace CCA-1 package builder.
 * Fills the project's compiled CCA-1 template for one subcontract, inserts the
 * approved Appendix A before Appendix B, and returns the merged PDF.
 * Runs in the Word add-in (browser) and in Node for testing. Requires pdf-lib.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('pdf-lib'));
  else root.DWPackage = factory(root.PDFLib);
})(typeof self !== 'undefined' ? self : this, function (PDFLib) {
  const { PDFDocument, StandardFonts } = PDFLib;

  // CCA 1-2021 is 29 pages (cover to signature page). Appendix A goes after it, before Appendix B.
  const DEFAULT_CCA_PAGES = 29;

  const ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
    'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
    'September', 'October', 'November', 'December'];

  function under1000(n) {
    const out = [];
    if (n >= 100) { out.push(ONES[Math.floor(n / 100)] + ' Hundred'); n %= 100; }
    if (n >= 20) { out.push(TENS[Math.floor(n / 10)] + (n % 10 ? '-' + ONES[n % 10] : '')); }
    else if (n > 0) { out.push(ONES[n]); }
    return out.join(' ');
  }

  /** 123456.78 -> "One Hundred Twenty-Three Thousand Four Hundred Fifty-Six and 78" (form prints "/100 dollars"). */
  function amountInWords(amount) {
    const cents = Math.round(amount * 100);
    let dollars = Math.floor(cents / 100);
    const c = String(cents % 100).padStart(2, '0');
    if (dollars === 0) return 'Zero and ' + c;
    const scales = ['', ' Thousand', ' Million', ' Billion'];
    const parts = [];
    let i = 0;
    while (dollars > 0) {
      const chunk = dollars % 1000;
      if (chunk) parts.unshift(under1000(chunk) + scales[i]);
      dollars = Math.floor(dollars / 1000);
      i++;
    }
    return parts.join(' ') + ' and ' + c;
  }

  function money(n) {
    return n.toLocaleString('en-CA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function ordinal(d) {
    const s = (d % 100 >= 11 && d % 100 <= 13) ? 'th' : ({ 1: 'st', 2: 'nd', 3: 'rd' }[d % 10] || 'th');
    return d + s;
  }

  /** Accepts a Date or "YYYY-MM-DD"; returns {day:"24th", month:"September", year:"2026"} or null. */
  function dateParts(v) {
    if (!v) return null;
    let d = v;
    if (typeof v === 'string') {
      const m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (!m) return null;
      d = new Date(+m[1], +m[2] - 1, +m[3]);
    }
    return { day: ordinal(d.getDate()), month: MONTHS[d.getMonth()], year: String(d.getFullYear()) };
  }

  /** Works out every value that goes into the CCA-1 from the subcontract record. */
  function computeValues(sub, gstRate) {
    const price = Number(sub.contractPrice);
    const gst = Math.round(price * gstRate) / 100;
    const total = Math.round((price + gst) * 100) / 100;
    return { price, gst, total };
  }

  /**
   * sub = {
   *   subcontractNumber, trade, subName, subAddress, subEmail, signerName, signerTitle,
   *   contractPrice (number), commenceDate, substantialDate (YYYY-MM-DD),
   *   pmName, pmEmail, agreementDate (YYYY-MM-DD, defaults to today)
   * }
   * Returns { pdfBytes, filled: {field: value}, warnings: [] }
   */
  async function buildPackage({ ccaTemplateBytes, appendixAPdfBytes, sub, ccaPageCount = DEFAULT_CCA_PAGES, lockFilledFields = true }) {
    const warnings = [];
    const doc = await PDFDocument.load(ccaTemplateBytes);
    const form = doc.getForm();
    const font = await doc.embedFont(StandardFonts.Helvetica);

    const has = (name) => { try { form.getField(name); return true; } catch (e) { return false; } };
    const filled = {};
    function set(name, value) {
      if (value === undefined || value === null || value === '') return;
      if (!has(name)) { warnings.push('Field not found in template: ' + name); return; }
      const f = form.getTextField(name);
      const text = String(value);
      // The template uses auto-size fonts, which render huge. Pick a size that fits the box.
      const r = f.acroField.getWidgets()[0].getRectangle();
      const fitWidth = (r.width - 6) / Math.max(font.widthOfTextAtSize(text, 1), 1);
      let size = Math.min(11, fitWidth, r.height * 0.7);
      if (size < 7 && r.height >= 24) {
        // Tall box with long text (e.g. price in words): wrap onto two lines.
        f.enableMultiline();
        size = Math.min(10, Math.max(7, fitWidth * 2.1), r.height / 2.6);
      }
      f.setText(text);
      f.setFontSize(Math.max(6, Math.floor(size * 2) / 2));
      if (lockFilledFields) f.enableReadOnly();
      filled[name] = text;
    }

    // GST rate is set once per project in the template (Article 5.2). Default 5%.
    let gstRate = 5;
    if (has('5.2 value added taxes')) {
      const r = parseFloat(form.getTextField('5.2 value added taxes').getText());
      if (!isNaN(r)) gstRate = r;
    }
    const v = computeValues(sub, gstRate);
    const agreed = dateParts(sub.agreementDate || new Date());
    const start = dateParts(sub.commenceDate);
    const finish = dateParts(sub.substantialDate);

    set('subcontract work', [sub.subcontractNumber, (sub.trade || '').toUpperCase()].filter(Boolean).join(' '));
    set('day', agreed.day); set('month', agreed.month); set('year', agreed.year);
    set('subcontractor', sub.subName);
    if (start) { set('day2', start.day); set('month2', start.month); set('year2', start.year); }
    else warnings.push('No commencement date on the Appendix A; Article 4 start date left blank.');
    if (finish) { set('day3', finish.day); set('month3', finish.month); set('year3', finish.year); }
    else warnings.push('No substantial completion date on the Appendix A; Article 4 completion date left blank.');
    set('5.1 subcontract price', amountInWords(v.price)); set('5.1price', money(v.price));
    set('5.2 value added price', amountInWords(v.gst)); set('5.2price', money(v.gst));
    set('5.3 total payable  price', amountInWords(v.total)); set('5.3price', money(v.total));
    set('subcontractor address', sub.subAddress);
    set('subcontractor email', sub.subEmail);
    set('contractor email', sub.pmEmail);
    set('witness name', sub.pmName ? sub.pmName + ', Project Manager' : '');
    set('subcontractor name and title', [sub.signerName, sub.signerTitle].filter(Boolean).join(', '));
    if (!sub.signerTitle) warnings.push('No signing authority title on the Appendix A; the trade will need to add it in DocuSign.');
    // Appendix E (statutory declaration) header and EFT form
    set('Subcontractor', sub.subName);
    set('Subcontract Number', sub.subcontractNumber);
    set('Subcontract Date', [agreed.month, agreed.day.replace(/\D+$/, '') + ',', agreed.year].join(' '));
    set('Supplier Name', sub.subName);

    // Only fields we set are redrawn (pdf-lib updates dirty fields on save), so the CA's own entries keep their look.
    form.updateFieldAppearances(font);

    // Merge: CCA-1 pages, then Appendix A, then Appendices B onward.
    const out = await PDFDocument.create();
    const ccaSaved = await PDFDocument.load(await doc.save());
    const appA = await PDFDocument.load(appendixAPdfBytes);
    const total = ccaSaved.getPageCount();
    if (total <= ccaPageCount) warnings.push('Template has no pages after the CCA-1; Appendix A appended at the end.');

    // Copying pages keeps their widgets but drops the AcroForm link, so rebuild the form after.
    // Simpler and reliable: insert Appendix A pages into the filled document itself.
    const appAPages = await ccaSaved.copyPages(appA, appA.getPageIndices());
    appAPages.forEach((p, i) => ccaSaved.insertPage(Math.min(ccaPageCount, total) + i, p));
    ccaSaved.setTitle(`${sub.subcontractNumber} ${sub.subName} CCA-1 Subcontract`);
    const pdfBytes = await ccaSaved.save();
    return { pdfBytes, filled, warnings, values: v };
  }

  /** Suggested file name for the compiled package. */
  function packageFileName(sub) {
    const clean = (s) => String(s || '').replace(/[\\/:*?"<>|#%]+/g, '').trim();
    return `CCA1-2021_${clean(sub.subcontractNumber)}_${clean(sub.trade)}_${clean(sub.subName)}_Contract_Compiled.pdf`;
  }

  return { buildPackage, amountInWords, money, dateParts, computeValues, packageFileName };
});
