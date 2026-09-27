/*
 * Document channel: lets the panel work without an app registration.
 * Each action updates the file's own status and history right away and adds a request to a
 * queue in the DW Request column. Word saves the file, SharePoint copies the values into the
 * library columns, and a Power Automate flow processes any request it hasn't seen yet.
 * The panel never waits for the flow, so people can keep working with the file open.
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
    let lastErr = null;
    // Newest part last: after a rewrite the updated copy is the one to use.
    for (const part of (parts || []).slice().reverse()) {
      try {
        const xml = await cb((done) => part.getXmlAsync(done));
        const doc = new DOMParser().parseFromString(xml, "application/xml");
        const els = doc.getElementsByTagNameNS("*", "DWRequest");
        if (els.length) return { part, xml, doc, el: els[0] };
      } catch (e) { lastErr = e; }
    }
    if (lastErr) throw lastErr;
    return null;
  }

  const textOf = (doc, name) => {
    const els = doc.getElementsByTagNameNS("*", name);
    return els.length ? els[0].textContent || "" : "";
  };

  /** Returns {request, status, history} or null when the file has no DW columns.
   *  Word can take a few seconds after opening to load the SharePoint fields, so keep trying. */
  async function read(tries) {
    let p = null, err = null;
    for (let i = 0; i < (tries || 1); i++) {
      try { p = await findPart(); err = null; } catch (e) { err = e; }
      if (p) break;
      if (i < (tries || 1) - 1) await new Promise((r) => setTimeout(r, 1000));
    }
    if (err) throw err;
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
    // Fallback: add the updated copy first, then remove the old one, so the fields are never missing.
    await cb((done) => Office.context.document.customXmlParts.addAsync(newXml, done));
    await cb((done) => p.part.deleteAsync(done));
  }

  const MAX_AGE_MS = 3 * 24 * 3600 * 1000, MAX_QUEUE = 10;

  /** The DW Request column holds a list of recent requests (older panels saved a single one). */
  function parseQueue(text) {
    if (!text || !text.trim()) return [];
    try {
      const v = JSON.parse(text);
      return Array.isArray(v) ? v : [v];
    } catch (e) { return []; }
  }

  function setText(doc, name, value) {
    const els = doc.getElementsByTagNameNS("*", name);
    if (!els.length) throw new Error(`This file's library has no ${name} column.`);
    els[0].removeAttributeNS(XSI, "nil");
    els[0].textContent = value;
  }

  /**
   * Applies an action to the file straight away: adds the request to the queue that Power Automate reads,
   * sets the status and adds the history line. Then saves. Returns the new {status, history}.
   * Nothing waits on Power Automate, so the next action can follow right away.
   */
  async function apply(req, status, historyLine) {
    const p = await findPart();
    if (!p) throw new Error("This file isn't set up for approvals (no DW Request column in its library).");
    const now = Date.now();
    const queue = parseQueue(textOf(p.doc, "DWRequest"))
      .filter((r) => r && r.id && (!r.sentAt || now - Date.parse(r.sentAt) < MAX_AGE_MS));
    queue.push(req);
    const history = [textOf(p.doc, "DWHistory").trim(), historyLine].filter(Boolean).join("\n");
    setText(p.doc, "DWRequest", JSON.stringify(queue.slice(-MAX_QUEUE)));
    if (status != null) setText(p.doc, "DWStatus", status);
    setText(p.doc, "DWHistory", history);
    await replacePart(p, new XMLSerializer().serializeToString(p.doc));
    const check = await read();
    if (!check || check.request.indexOf(req.id) < 0) throw new Error("Word did not accept the change. Try again in a moment.");
    await Word.run(async (ctx) => { ctx.document.save(); await ctx.sync(); });
    return { status: check.status, history: check.history };
  }

  /** Name of whoever last saved the file, used as a first guess for the person's name. */
  async function lastAuthor() {
    try {
      return await Word.run(async (ctx) => {
        const props = ctx.document.properties;
        props.load("lastAuthor");
        await ctx.sync();
        return props.lastAuthor || "";
      });
    } catch (e) { return ""; }
  }

  window.DWDoc = { read, apply, parseQueue, lastAuthor };
})();
