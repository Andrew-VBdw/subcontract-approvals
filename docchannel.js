/*
 * Document channel: lets the panel work without an app registration.
 * Requests are written into the file's SharePoint properties (the DW Request column),
 * Word saves the file, SharePoint copies the value into the library column, and a
 * Power Automate flow picks it up. The flow writes status back into the DW Status and
 * DW History columns, which appear in the file the next time it is opened.
 */
(function () {
  const NS = "http://schemas.microsoft.com/office/2006/metadata/properties";
  const cb = (fn) => new Promise((res, rej) => fn((r) => (r.status === "succeeded" ? res(r.value) : rej(r.error || new Error("Office call failed")))));

  async function propsNodes() {
    const parts = await cb((done) => Office.context.document.customXmlParts.getByNamespaceAsync(NS, done));
    if (!parts || !parts.length) return null;
    const nodes = await cb((done) => parts[0].getNodesAsync("/*/*/*", done));
    const map = {};
    nodes.forEach((n) => { map[n.baseName] = n; });
    return { part: parts[0], nodes: map };
  }

  /** Returns {request, status, history} or null when the file has no DW columns. */
  async function read() {
    const p = await propsNodes();
    if (!p || !p.nodes.DWRequest) return null;
    const text = async (n) => (n ? (await cb((done) => n.getTextAsync(done))) || "" : "");
    return { request: await text(p.nodes.DWRequest), status: await text(p.nodes.DWStatus), history: await text(p.nodes.DWHistory) };
  }

  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  /** Writes the request into the file and saves it. */
  async function write(obj) {
    const p = await propsNodes();
    if (!p || !p.nodes.DWRequest) throw new Error("This file isn't set up for approvals (no DW Request column in its library).");
    const node = p.nodes.DWRequest;
    const json = JSON.stringify(obj);
    try {
      await cb((done) => node.setXmlAsync(`<DWRequest xmlns="${node.namespaceUri}">${esc(json)}</DWRequest>`, done));
    } catch (e) {
      await cb((done) => node.setTextAsync(json, done));
    }
    const check = await cb((done) => node.getTextAsync(done));
    if (!check || check.indexOf(obj.id) < 0) throw new Error("Word did not accept the request. Try again after the file finishes saving.");
    await Word.run(async (ctx) => { ctx.document.save(); await ctx.sync(); });
  }

  window.DWDoc = { read, write };
})();
