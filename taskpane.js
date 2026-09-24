/*
 * Dawson Wallace Subcontract Approvals: Word task pane.
 * Reads the open Appendix A, shows its status and history, and sends actions to the
 * hub site's Requests list, where Power Automate picks them up.
 */
(function () {
  const C = window.DW_CONFIG;
  const $ = (id) => document.getElementById(id);
  const PRE_EXECUTED = ["Draft", "Pending Approval", "Approved", "With Trade", "With Accounting",
    "With CA", "Out for Signature", "Signature Issue"];

  const state = { doc: null, req: null, docUrl: null, record: null, project: null, me: null };

  // ---------- helpers ----------
  function show(id, on) { $(id).hidden = !on; }
  function message(text, kind) {
    const m = $("msg");
    m.textContent = text; m.className = "msg " + (kind || "info"); m.hidden = !text;
  }
  function money(n) {
    return n == null || isNaN(n) ? "" : "$" + Number(n).toLocaleString("en-CA", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function parseMoney(s) {
    if (s == null) return NaN;
    const n = parseFloat(String(s).replace(/[^0-9.\-]/g, ""));
    return isNaN(n) ? NaN : n;
  }
  function isPlaceholder(t) { return !t || C.placeholderPatterns.some((re) => re.test(t.trim())); }
  function toIsoDate(text) {
    if (isPlaceholder(text)) return "";
    const m = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[0];
    const d = new Date(text);
    if (isNaN(d)) return "";
    return [d.getFullYear(), String(d.getMonth() + 1).padStart(2, "0"), String(d.getDate()).padStart(2, "0")].join("-");
  }
  function fieldList(el, rows) {
    el.innerHTML = "";
    rows.filter((r) => r[1] !== undefined && r[1] !== null && r[1] !== "").forEach(([k, v]) => {
      const dt = document.createElement("dt"); dt.textContent = k;
      const dd = document.createElement("dd"); dd.textContent = v;
      el.append(dt, dd);
    });
  }
  function badgeClass(status) {
    if (["Executed", "Compliance Complete", "Approved"].includes(status)) return "badge ok";
    if (["Signature Issue", "Cancelled"].includes(status)) return "badge bad";
    if (status === "Draft") return "badge";
    return "badge warn";
  }

  // ---------- read the Appendix A ----------
  async function readDocument() {
    return Word.run(async (ctx) => {
      const ccs = ctx.document.contentControls;
      ccs.load("items/tag,items/text");
      await ctx.sync();
      const byTag = {};
      ccs.items.forEach((cc) => { if (cc.tag) byTag[cc.tag] = (cc.text || "").trim(); });
      const t = C.tags, r = C.requirementTags;
      const val = (tag) => (isPlaceholder(byTag[tag]) ? "" : byTag[tag] || "");
      const checked = (tag) => byTag[tag] === "☒"; // ☒
      const doc = {
        found: Object.keys(byTag).some((k) => k.startsWith("DW_")),
        trade: val(t.trade),
        projectName: val(t.projectName),
        projectNumber: val(t.projectNumber),
        projectAddress: val(t.projectAddress),
        subcontractNumber: val(t.subcontractNumber),
        subName: val(t.subName),
        subAddress: val(t.subAddress),
        signerName: val(t.signerName),
        signerTitle: val(t.signerTitle),
        subEmail: val(t.subEmail),
        contractPrice: parseMoney(val(t.contractPrice)),
        contractPriceWords: val(t.contractPriceWords),
        commenceDate: toIsoDate(byTag[t.commenceDate] || ""),
        substantialDate: toIsoDate(byTag[t.substantialDate] || ""),
        requirements: {
          bonds: checked(r.bonds),
          glLimit: checked(r.gl10m) ? "$10M" : checked(r.gl5m) ? "$5M" : "None",
          profLiab: checked(r.profLiab),
          wcb: checked(r.wcb)
        }
      };
      return doc;
    });
  }

  function docProblems(d) {
    const p = [];
    if (!d.found) { p.push("This document has no labelled Appendix A fields. Use the current Appendix A template."); return p; }
    if (!d.subcontractNumber) p.push("Subcontract number is blank.");
    if (!d.subName) p.push("Subcontractor legal name is blank.");
    if (!d.subEmail) p.push("Subcontractor email is blank.");
    if (isNaN(d.contractPrice)) p.push("Contract base price is blank or not a number.");
    if (!d.signerName) p.push("Signing authority is blank.");
    return p;
  }

  function renderDoc() {
    const d = state.doc;
    fieldList($("docFields"), [
      ["Subcontract", d.subcontractNumber], ["Trade", d.trade], ["Subcontractor", d.subName],
      ["Signing authority", [d.signerName, d.signerTitle].filter(Boolean).join(", ")], ["Email", d.subEmail],
      ["Contract price", isNaN(d.contractPrice) ? "" : money(d.contractPrice)],
      ["Commencement", d.commenceDate], ["Substantial completion", d.substantialDate],
      ["GL insurance", d.requirements.glLimit], ["Professional liability", d.requirements.profLiab ? "Required" : ""],
      ["Bonds", d.requirements.bonds ? "Required" : ""]
    ]);
    const w = $("docWarnings"); w.innerHTML = "";
    docProblems(d).forEach((t) => { const p = document.createElement("div"); p.className = "warnline"; p.textContent = t; w.append(p); });
    show("docSection", true);
  }

  // ---------- hub data ----------
  async function loadRecord() {
    const num = state.doc.subcontractNumber;
    state.record = null; state.project = null;
    if (num) {
      const r = await DWGraph.items("subcontracts", `fields/Title eq '${DWGraph.esc(num)}'`);
      state.record = r[0] || null;
    }
    if (state.doc.projectNumber) {
      const p = await DWGraph.items("projects", `fields/ProjectNumber eq '${DWGraph.esc(state.doc.projectNumber)}'`);
      state.project = p[0] || null;
    }
  }

  async function loadLog() {
    const list = $("log"); list.innerHTML = "";
    if (!state.record) { show("logSection", false); return; }
    const rows = await DWGraph.items("log", `fields/SubcontractNumber eq '${DWGraph.esc(state.record.Title)}'`, "fields/EventTime desc", 100);
    rows.forEach((r) => {
      const li = document.createElement("li");
      const when = r.EventTime ? new Date(r.EventTime).toLocaleString("en-CA", { dateStyle: "medium", timeStyle: "short" }) : "";
      li.innerHTML = `<div class="what"></div><div class="when"></div><div class="note"></div>`;
      li.querySelector(".what").textContent = r.EventType || r.Title;
      li.querySelector(".when").textContent = [when, r.ActorName, r.FileVersion ? "v" + r.FileVersion : ""].filter(Boolean).join(" · ");
      li.querySelector(".note").textContent = r.Comment || "";
      list.append(li);
    });
    show("logSection", rows.length > 0);
  }

  function renderStatus() {
    const r = state.record;
    const status = r ? r.Status : "Not submitted";
    $("statusBadge").textContent = status;
    $("statusBadge").className = r ? badgeClass(status) : "badge";
    $("statusNote").textContent = r ? "" : "This Appendix A has not been submitted yet.";
    fieldList($("recordFields"), r ? [
      ["Estimate carried", money(r.EstimateCarried)], ["Contract value", money(r.ContractValue)],
      ["Delta", r.Delta != null ? money(r.Delta) : ""], ["Cost codes", r.CostCodes],
      ["Approved version", r.ApprovedVersion], ["Envelope", r.EnvelopeId], ["Issue", r.IssueType]
    ] : []);
    show("statusSection", true);

    const canSubmit = !r || r.Status === "Draft";
    $("submitTitle").textContent = r ? "Resubmit for approval" : "Submit for approval";
    if (r && canSubmit) {
      if (r.EstimateCarried != null && !$("estimate").value) $("estimate").value = Number(r.EstimateCarried).toFixed(2);
      if (r.CostCodes && !$("costCodes").value) $("costCodes").value = r.CostCodes;
    }
    show("submitSection", canSubmit && docProblems(state.doc).length === 0);
    updateDelta();
    renderActions();
    show("commentSection", !!r);
  }

  function updateDelta() {
    const est = parseMoney($("estimate").value);
    const price = state.doc ? state.doc.contractPrice : NaN;
    const line = $("deltaLine");
    if (isNaN(est) || isNaN(price)) { line.hidden = true; show("justLabel", false); return; }
    const delta = Math.round((price - est) * 100) / 100;
    line.hidden = false;
    line.className = "delta " + (delta > 0 ? "over" : "under");
    line.textContent = delta > 0 ? `Over estimate by ${money(delta)}` : `Within estimate by ${money(-delta)}`;
    show("justLabel", delta > 0);
  }

  // ---------- actions ----------
  const ACTIONS = {
    SendToTrade: { label: "Send to Trade", statuses: ["Approved"], prompt: "Email a PDF of the approved Appendix A to " },
    TradeAlreadyAgreed: { label: "Trade Already Agreed", statuses: ["Approved"], prompt: "Skip the trade step. Attach the email where the trade agreed.", file: true, fileRequired: true },
    TradeAccepted: { label: "Trade Accepted", statuses: ["With Trade"], prompt: "Attach the trade's acceptance email.", file: true, fileRequired: true },
    TradeWantsChanges: { label: "Trade Wants Changes", statuses: ["With Trade"], prompt: "Returns to Draft. Describe what the trade wants changed.", noteRequired: true, file: true },
    BuildPackage: { label: "Build CCA Package", statuses: ["With CA"], prompt: "Fill the project CCA-1, insert this Appendix A before Appendix B, and create the DocuSign draft." },
    LogMarkup: { label: "Log Markup", statuses: ["Out for Signature"], prompt: "Attach the trade's markup. The open envelope will be voided.", file: true, fileRequired: true },
    LogDispute: { label: "Log Dispute", statuses: ["Out for Signature"], prompt: "Attach the email or summarize the call. The open envelope will be voided.", file: true, noteRequired: true },
    Resolve: { label: "Resolve Issue", statuses: ["Signature Issue"], prompt: "What was agreed, and where does it go next?", noteRequired: true,
      choices: ["Reissue unchanged", "Appendix A must change", "Cancel subcontract"] },
    Withdraw: { label: "Withdraw to Draft", statuses: PRE_EXECUTED.filter((s) => s !== "Draft"), prompt: "Pulls this back to Draft for your changes. It will need approval again.", noteRequired: true },
    Cancel: { label: "Cancel Subcontract", statuses: PRE_EXECUTED, prompt: "This subcontract is not proceeding. Any open envelope will be voided.", noteRequired: true }
  };
  let current = null;

  function renderActions() {
    const box = $("actions"); box.innerHTML = "";
    const r = state.record;
    if (!r) { show("actionSection", false); return; }
    Object.entries(ACTIONS).filter(([, a]) => a.statuses.includes(r.Status)).forEach(([key, a]) => {
      const b = document.createElement("button"); b.textContent = a.label;
      b.onclick = () => openAction(key); box.append(b);
    });
    $("actionForm").hidden = true;
    show("actionSection", box.children.length > 0);
  }

  function openAction(key) {
    const a = ACTIONS[key]; current = key;
    $("actionPrompt").textContent = a.prompt + (key === "SendToTrade" ? (state.doc.subEmail || "the subcontractor") + "." : "");
    $("actionNote").value = ""; $("actionFile").value = "";
    $("actionNote").hidden = key === "SendToTrade" || key === "BuildPackage";
    $("fileLabel").hidden = !a.file;
    const sel = $("actionChoice"); sel.innerHTML = ""; sel.hidden = !a.choices;
    (a.choices || []).forEach((c) => { const o = document.createElement("option"); o.textContent = c; sel.append(o); });
    $("actionForm").hidden = false;
  }

  async function uploadAttachment(file) {
    if (!file) return "";
    const di = await DWGraph.driveItemFromUrl(state.docUrl);
    const folder = (state.project && state.project.CorrespondenceFolder) || "Subcontracts/Correspondence";
    const bytes = new Uint8Array(await file.arrayBuffer());
    const name = `${state.record.Title} ${new Date().toISOString().slice(0, 10)} ${file.name}`;
    const up = await DWGraph.upload(di.parentReference.driveId, folder, name, bytes);
    return up.webUrl;
  }

  async function sendRequest(action, payload) {
    await DWGraph.createItem("requests", {
      Title: action,
      SubcontractNumber: state.doc.subcontractNumber,
      DocumentUrl: state.docUrl,
      Payload: JSON.stringify(payload || {}),
      RequestStatus: "New"
    });
  }

  async function runAction() {
    const a = ACTIONS[current];
    const note = $("actionNote").value.trim();
    const file = $("actionFile").files[0];
    if (a.noteRequired && !note && !(current === "LogDispute" && file)) return message("Please add a note.", "error");
    if (a.fileRequired && !file) return message("Please attach the email or file.", "error");
    $("actionGo").disabled = true;
    try {
      if (current === "BuildPackage") { await buildPackage(); }
      else {
        message("Working…", "info");
        const link = await uploadAttachment(file);
        await sendRequest(current, { note, attachmentUrl: link, choice: a.choices ? $("actionChoice").value : "" });
        message(`${a.label} sent. The status will update in a minute.`, "ok");
      }
      $("actionForm").hidden = true;
      pollForChange();
    } catch (e) {
      message("Could not complete: " + e.message, "error");
    } finally { $("actionGo").disabled = false; }
  }

  async function buildPackage() {
    if (!state.project || !state.project.CCATemplateUrl) throw new Error("The project has no CCA-1 template set in the Projects list.");
    message("Building the CCA-1 package…", "info");
    const tplUrl = state.project.CCATemplateUrl.Url || state.project.CCATemplateUrl;
    const tpl = await DWGraph.driveItemFromUrl(tplUrl);
    const tplBytes = await DWGraph.fileBytes(tpl.parentReference.driveId, tpl.id);
    const di = await DWGraph.driveItemFromUrl(state.docUrl);
    const appA = await DWGraph.fileAsPdf(di.parentReference.driveId, di.id);
    const r = state.record, d = state.doc;
    const sub = {
      subcontractNumber: d.subcontractNumber, trade: d.trade, subName: d.subName, subAddress: d.subAddress,
      subEmail: d.subEmail, signerName: d.signerName, signerTitle: d.signerTitle,
      contractPrice: d.contractPrice, commenceDate: d.commenceDate, substantialDate: d.substantialDate,
      pmName: r.PMName || "", pmEmail: r.PMEmail || ""
    };
    const out = await DWPackage.buildPackage({ ccaTemplateBytes: tplBytes, appendixAPdfBytes: appA, sub, ccaPageCount: C.ccaPageCount });
    const folder = state.project.PackagesFolder || "Subcontracts/Packages";
    const up = await DWGraph.upload(di.parentReference.driveId, folder, DWPackage.packageFileName(sub), out.pdfBytes);
    await sendRequest("PackageBuilt", { packageUrl: up.webUrl, packageDriveId: di.parentReference.driveId, packageItemId: up.id, warnings: out.warnings });
    message("Package saved. The DocuSign draft will be ready for review shortly." +
      (out.warnings.length ? " Check: " + out.warnings.join(" ") : ""), out.warnings.length ? "info" : "ok");
  }

  async function submit() {
    const d = state.doc;
    const est = parseMoney($("estimate").value);
    const codes = $("costCodes").value.trim();
    const just = $("justification").value.trim();
    if (isNaN(est)) return message("Enter the estimate carried.", "error");
    if (!codes) return message("Enter the cost code(s).", "error");
    const delta = Math.round((d.contractPrice - est) * 100) / 100;
    if (delta > 0 && !just) return message("Explain why the contract is over the estimate.", "error");
    const gst = Math.round(d.contractPrice * 5) / 100;
    $("submitBtn").disabled = true;
    try {
      message("Submitting…", "info");
      await sendRequest(state.record ? "Resubmit" : "Submit", {
        estimateCarried: est, costCodes: codes, justification: just,
        contractValue: d.contractPrice, delta,
        doc: d,
        words: {
          price: DWPackage.amountInWords(d.contractPrice),
          gst: DWPackage.amountInWords(gst),
          total: DWPackage.amountInWords(d.contractPrice + gst)
        }
      });
      message("Submitted. The approver has been notified, and you'll get a confirmation email.", "ok");
      show("submitSection", false);
      pollForChange();
    } catch (e) {
      message("Could not submit: " + e.message, "error");
    } finally { $("submitBtn").disabled = false; }
  }

  async function addComment() {
    const text = $("comment").value.trim();
    if (!text) return;
    try {
      await sendRequest("Comment", { note: text });
      $("comment").value = "";
      message("Comment added.", "ok");
      setTimeout(refresh, 8000);
    } catch (e) { message("Could not add comment: " + e.message, "error"); }
  }

  let pollTimer = null;
  function pollForChange() {
    const before = state.record ? state.record.Status + "|" + state.record.Modified : "none";
    let tries = 0;
    clearInterval(pollTimer);
    pollTimer = setInterval(async () => {
      tries++;
      try {
        await loadRecord();
        const now = state.record ? state.record.Status + "|" + state.record.Modified : "none";
        if (now !== before || tries > 12) { clearInterval(pollTimer); renderStatus(); await loadLog(); }
      } catch (e) { clearInterval(pollTimer); }
    }, 5000);
  }

  async function refresh() {
    await loadRecord();
    renderStatus();
    await loadLog();
  }

  async function start() {
    try {
      state.docUrl = Office.context.document.url || "";
      state.doc = await readDocument();
      show("loading", false);
      renderDoc();
      if (!/^https:\/\//i.test(state.docUrl)) {
        message("Open this Appendix A from SharePoint (or with AutoSave on) so it can be linked to the workflow.", "error");
        return;
      }
      await refresh();
    } catch (e) {
      show("loading", false);
      message("Could not load: " + e.message, "error");
    }
  }

  Office.onReady(() => {
    $("estimate").addEventListener("input", updateDelta);
    $("submitBtn").onclick = submit;
    $("actionGo").onclick = runAction;
    $("actionCancel").onclick = () => { $("actionForm").hidden = true; };
    $("commentBtn").onclick = addComment;
    start();
  });
})();
