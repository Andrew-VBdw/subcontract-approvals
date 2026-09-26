/*
 * Document channel: lets the panel work without an app registration.
 * Requests are written into the file's SharePoint properties (the DW Request column),
 * Word saves the file, SharePoint copies the value into the library column, and a
 * Power Automate flow picks it up. The flow writes status back into the DW Status and
 * DW History columns, which appear in the file the next time it is opened.
 *
 * The whole properties block is read and rewritten at once. Editing single nodes
 * fails in Word when SharePoint stored the field as empty ("The specified node was not found").
 */
(function () {
  const NS = "http://schemas.microsoft.com/office/2006/metadata/properties";
  const XSI = "http://www.w3.org/2001/XMLSchema-instance";
  const cb = (fn) => new Promise((res, rej) => fn((r) => (r.status === "succeeded" ? res(r.value) : rej(r.error || new Error("Office call failed")))));

  /** Finds the properties part that holds the DW columns. Returns {part, xml, doc, el} or null. */
  async function findPart() {
    const parts = await cb((done) => Office.context.document.customXmlParts.getByNamespaceAsync(NS, done));
    for (const part of parts || []) {
      const xml = await cb((done) => part.getXmlAsync(done));
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      const els = doc.getElementsByTagNameNS("*", "DWRequest");
      if (els.length) return { part, xml, doc, el: els[0] };
    }
    return null;
  }

  const textOf = (doc, name) => {
    const els = doc.getElementsByTagNameNS("*", name);
    return els.length ? els[0].textContent || "" : "";
  };

  /** Returns {request, status, history} or null when the file has no DW columns. */
  async function read() {
    const p = await findPart();
    if (!p) return null;
    return { request: textOf(p.doc, "DWRequest"), status: textOf(p.doc, "DWStatus"), history: textOf(p.doc, "DWHistory") };
  }

  async function replacePart(p, newXml) {
    // Preferred: rewrite the same part in place (Word API 1.4).
    if (window.Word && Office.context.requirements.isSetSupported("WordApi", "1.4")) {
      try {
        await Word.run(async (ctx) => {
          const parts = ctx.document.customXmlParts.getByNamespace(NS);
          parts.load("items/id");
          await ctx.sync();
          const id = p.part.id;
          const target = parts.items.find((x) => x.id === id) || parts.items[0];
          target.setXml(newXml);
          await ctx.sync();
        });
        return;
      } catch (e) {
        console.warn("setXml failed, falling back to replace", e);
      }
    }
    // Fallback: remove the part and add the updated copy.
    await cb((done) => p.part.deleteAsync(done));
    await cb((done) => Office.context.document.customXmlParts.addAsync(newXml, done));
  }

  /** Writes the request into the file and saves it. */
  async function write(obj) {
    const p = await findPart();
    if (!p) throw new Error("This file isn't set up for approvals (no DW Request column in its library).");
    p.el.removeAttributeNS(XSI, "nil");
    p.el.textContent = JSON.stringify(obj);
    const newXml = new XMLSerializer().serializeToString(p.doc);
    await replacePart(p, newXml);
    const check = await read();
    if (!check || check.request.indexOf(obj.id) < 0) throw new Error("Word did not accept the request. Try again after the file finishes saving.");
    await Word.run(async (ctx) => { ctx.document.save(); await ctx.sync(); });
  }

  window.DWDoc = { read, write };
})();
