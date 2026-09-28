/*
 * Dawson Wallace Subcontract Approvals: Word task pane.
 * Reads the open Appendix A, shows its status and history, and sends actions to the
 * hub site's Requests list, where Power Automate picks them up.
 */
(function () {
  const C = window.DW_CONFIG;
  const $ = (id) => document.getElementById(id);
  const PRE_EXECUTED = ["Draft", "Pending Approval", "Approved", "With Trade", "With Accounting", "Returned by Accounting",
    "Ready to Issue", "Out for Signature", "Signature Issue"];

  const state = { doc: null, req: null, docUrl: null, record: null, project: null, region: null, me: null };

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
      const dd = document.createElement("dd");
      if (v && typeof v === "object" && v.link) { const a = document.createElement("a"); a.href = v.link; a.target = "_blank"; a.textContent = v.text; dd.append(a); }
      else dd.textContent = v;
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
      // Primary source: the document's own XML. It covers every field type, including checkboxes
      // and date pickers inside table cells, which the content-control API does not report reliably.
      const ooxml = ctx.document.body.getOoxml();
      const ccs = ctx.document.contentControls;
      ccs.load("items/tag,items/text");
      await ctx.sync();
      const byTag = {}, checks = {}, dates = {};
      try { readOoxml(ooxml.value, byTag, checks, dates); } catch (e) { /* fall back to the API below */ }
      ccs.items.forEach((cc) => { if (cc.tag && !(cc.tag in byTag)) byTag[cc.tag] = (cc.text || "").trim(); });
      const t = C.tags, r = C.requirementTags;
      const val = (tag) => (isPlaceholder(byTag[tag]) ? "" : byTag[tag] || "");
      const checked = (tag) => (tag in checks ? checks[tag] : byTag[tag] === "☒");
      const dateOf = (tag) => dates[tag] || toIsoDate(byTag[tag] || "");
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
        commenceDate: dateOf(t.commenceDate),
        substantialDate: dateOf(t.substantialDate),
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

  /** Reads every tagged content control from the document XML: text, checkbox state and date value. */
  function readOoxml(xml, byTag, checks, dates) {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    const dom = new DOMParser().parseFromString(xml, "application/xml");
    const sdts = dom.getElementsByTagNameNS(W, "sdt");
    for (const sdt of Array.from(sdts)) {
      const pr = Array.from(sdt.children).find((c) => c.localName === "sdtPr");
      if (!pr) continue;
      const tagEl = Array.from(pr.children).find((c) => c.localName === "tag");
      const tag = tagEl && (tagEl.getAttributeNS(W, "val") || tagEl.getAttribute("w:val"));
      if (!tag) continue;
      const content = Array.from(sdt.children).find((c) => c.localName === "sdtContent");
      const text = content ? Array.from(content.getElementsByTagNameNS(W, "t")).map((n) => n.textContent).join("").trim() : "";
      byTag[tag] = text;
      const cb = pr.getElementsByTagNameNS(W14, "checkbox")[0];
      if (cb) {
        const ch = cb.getElementsByTagNameNS(W14, "checked")[0];
        const v = ch && (ch.getAttributeNS(W14, "val") || ch.getAttribute("w14:val"));
        checks[tag] = v === "1" || v === "true" || (!ch && text === "☒");
        if (!ch) checks[tag] = text === "☒";
      }
      const dt = Array.from(pr.children).find((c) => c.localName === "date");
      if (dt) {
        const full = dt.getAttributeNS(W, "fullDate") || dt.getAttribute("w:fullDate");
        if (full) dates[tag] = full.slice(0, 10);
      }
    }
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
    state.region = null;
    if (state.project && state.project.RegionLookupId) {
      const g = await DWGraph.items("regions");
      state.region = g.find((x) => String(x._id) === String(state.project.RegionLookupId)) || null;
    }
    if (!state.settings) {
      try { state.settings = Object.fromEntries((await DWGraph.items("settings")).map((i) => [i.Title, i.SettingValue])); }
      catch (e) { state.settings = {}; }
    }
    state.test = /^on$/i.test(state.settings.TestMode || "");
    if (!state.me) { try { state.me = await DWGraph.me(); } catch (e) { /* role buttons stay hidden */ } }
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
    $("statusNote").textContent = r ? (NEXT[r.Status] || "") : NOT_SUBMITTED;
    fieldList($("recordFields"), r ? [
      ["Estimate carried", money(r.EstimateCarried)], ["Contract value", money(r.ContractValue)],
      ["Delta", r.Delta != null ? money(r.Delta) : ""], ["Cost codes", r.CostCodes],
      ["Approved version", r.ApprovedVersion], ["CCA-1 package", r.PackageUrl ? { link: r.PackageUrl.Url || r.PackageUrl, text: "Open PDF" } : ""],
      ["Envelope", r.EnvelopeId], ["Issue", r.IssueType]
    ] : []);
    show("statusSection", true);

    const canSubmit = !r || r.Status === "Draft";
    $("submitTitle").textContent = r ? "Resubmit for approval" : "Submit for approval";
    show("backSection", !!r && r.Status === "Returned by Accounting");
    if (r && canSubmit) {
      if (r.EstimateCarried != null && !$("estimate").value) $("estimate").value = Number(r.EstimateCarried).toFixed(2);
      if (r.CostCodes && !$("costCodes").value) $("costCodes").value = r.CostCodes;
    }
    const projectReady = state.channel === "doc" || (!!state.project && state.project.Active !== false);
    show("submitSection", canSubmit && docProblems(state.doc).length === 0 && (projectReady || state.preview));
    if (!state.preview && !projectReady && !r) {
      $("statusNote").textContent = `Project ${state.doc.projectNumber || "(blank)"} isn't set up for approvals yet. The pilot is running on 26-205 MAPEI only.`;
    }
    updateDelta();
    renderActions();
    show("commentSection", !!r);
  }

  function updateDelta() {
    const est = parseMoney($("estimate").value);
    const price = state.doc ? state.doc.contractPrice : NaN;
    const line = $("deltaLine");
    if (isNaN(est) || isNaN(price)) { line.hidden = true; return; }
    const delta = Math.round((price - est) * 100) / 100;
    line.hidden = false;
    line.className = "delta " + (delta > 0 ? "over" : "under");
    line.textContent = delta > 0 ? `Over estimate by ${money(delta)}` : `Within estimate by ${money(-delta)}`;
  }

  // ---------- actions ----------
  // role: "approver" and "accounting" buttons only show for that person (the approver sees everything).
  // Everything else shows for anyone working on the file; the log records who clicked.
  const ACTIONS = {
    Approve: { label: "Approve", statuses: ["Pending Approval"], role: "approver", prompt: "Moves it to Approved and emails the PM to send it to the trade. Any older approval email for it stops working. Add a note if you like." },
    ReturnToDraft: { label: "Return to Draft", statuses: ["Pending Approval"], role: "approver", prompt: "Returns it to Draft for the PM to revise and resubmit. Any older approval email for it stops working. What needs to change?", noteRequired: true },
    SendToTrade: { label: "Send to Trade", statuses: ["Approved"], prompt: "Email the approved Appendix A (Word file) to " },
    TradeAlreadyAgreed: { label: "Trade Already Agreed", statuses: ["Approved"], prompt: "Skips sending it to the trade and moves it to With Accounting. Accounting gets an email to enter it. Add a note if you like." },
    TradeAccepted: { label: "Trade Accepted", statuses: ["With Trade"], prompt: "Moves it to With Accounting. Accounting gets an email to enter it. Add a note if you like." },
    TradeWantsChanges: { label: "Trade Wants Changes", statuses: ["With Trade"], prompt: "Returns it to Draft for the PM to revise and resubmit. Add a note if you like." },
    AccountingComplete: { label: "Entered into Accounting", statuses: ["With Accounting"], role: "accounting", prompt: "Moves it to Ready to Issue and emails the PM to build the CCA-1 package. Any older accounting email for it stops working. If accounting changed the subcontract number or cost codes, correct them here. A new number is also updated in this Appendix A.", codes: true, number: true },
    ReturnFromAccounting: { label: "Send Back to PM", statuses: ["With Accounting"], role: "accounting", prompt: "Returns it to the PM to fix. The PM can send it straight back to accounting without another approval. What does the PM need to fix?", noteRequired: true },
    BackToAccounting: { label: "Back to Accounting", statuses: ["Returned by Accounting"], prompt: "Sends it straight back to accounting. No new approval needed. Accounting gets an email to enter it. Say what you fixed." },
    BuildPackage: { needsGraph: true, label: "Build CCA-1 Package", statuses: ["Ready to Issue"], prompt: "Fill the project CCA-1 and insert this Appendix A before Appendix B. You can review the PDF before sending." },
    DraftAndSend: { label: "Draft Contract & Send to DocuSign", statuses: ["Ready to Issue"], comingSoon: true, prompt: "" },
    SendForSignature: { label: "Send for Signature", statuses: ["Ready to Issue"], needsPackage: true, prompt: "Send the package through DocuSign to the subcontractor's signer, then the Dawson Wallace signer." },
    LogMarkup: { label: "Log Markup", statuses: ["Out for Signature"], prompt: "Attach the trade's markup. The open envelope will be voided.", file: true, fileRequired: true },
    LogDispute: { label: "Log Dispute", statuses: ["Out for Signature"], prompt: "Attach the email or summarize the call. The open envelope will be voided.", file: true, noteRequired: true },
    Resolve: { label: "Resolve Issue", statuses: ["Signature Issue"], prompt: "What was agreed, and where does it go next?", noteRequired: true,
      choices: ["Reissue unchanged", "Appendix A must change", "Cancel subcontract"] },
    Withdraw: { label: "Withdraw to Draft", statuses: PRE_EXECUTED.filter((s) => s !== "Draft"), prompt: "Pulls this back to Draft for your changes. Any open envelope is voided and it will need approval again. Logged with your note.", noteRequired: true },
    Cancel: { label: "Cancel Subcontract", statuses: PRE_EXECUTED, prompt: "This subcontract is not proceeding. Any open envelope will be voided. Logged with your note.", noteRequired: true }
  };
  const myEmail = () => ((state.me && (state.me.mail || state.me.userPrincipalName)) || "").toLowerCase();
  function allowed(a) {
    if (!a.role || state.preview || state.channel === "doc") return true;
    const me = myEmail(), g = state.region || {};
    const approver = (g.ApproverEmail || "").toLowerCase(), acct = (g.AccountingEmail || "").toLowerCase();
    if (me && me === approver) return true;
    if (a.role === "accounting") return me && me === acct;
    return false;
  }
  let current = null;

  // Every step's buttons are always listed in workflow order; only the ones for the current stage
  // (and the person signed in) are clickable. The rest are greyed with a note on when they unlock.
  const STEPS = [
    { title: "1. Approval", statuses: ["Pending Approval"], keys: ["Approve", "ReturnToDraft"], who: "Approver only" },
    { title: "2. Trade", statuses: ["Approved", "With Trade"], keys: ["SendToTrade", "TradeAlreadyAgreed", "TradeAccepted", "TradeWantsChanges"] },
    { title: "3. Accounting", statuses: ["With Accounting"], keys: ["AccountingComplete", "ReturnFromAccounting"], who: "Accounting only" },
    { title: "4. Contract & DocuSign", statuses: ["Ready to Issue"], keys: ["DraftAndSend"] },
    { title: "5. Signature", statuses: ["Out for Signature", "Signature Issue"], keys: ["LogMarkup", "LogDispute", "Resolve"] },
    { title: "Any time before signing", statuses: [], keys: ["Withdraw", "Cancel"] }
  ];
  const WHEN = {
    Approve: "After the PM submits", ReturnToDraft: "After the PM submits",
    SendToTrade: "After approval", TradeAlreadyAgreed: "After approval",
    TradeAccepted: "After it's sent to the trade", TradeWantsChanges: "After it's sent to the trade",
    AccountingComplete: "After the trade accepts", ReturnFromAccounting: "After the trade accepts", BackToAccounting: "If accounting sends it back",
    BuildPackage: "After it's entered into accounting", SendForSignature: "After the package is built",
    LogMarkup: "While out for signature", LogDispute: "While out for signature", Resolve: "If there's a signature issue",
    Withdraw: "After submitting, until signed", Cancel: "Until signed"
  };

  function renderActions() {
    const box = $("actions");
    $("actionSection").append($("actionForm"));
    box.innerHTML = "";
    const r = state.record;
    const status = r ? r.Status : "Not submitted";
    STEPS.forEach((step) => {
      const g = document.createElement("div");
      g.className = "step" + (step.statuses.includes(status) ? " current" : "");
      const h = document.createElement("div"); h.className = "step-title"; h.textContent = step.title;
      g.append(h);
      step.keys.forEach((key) => {
        const a = ACTIONS[key];
        const stageOk = !!r && a.statuses.includes(status) && (!a.needsPackage || r.PackageUrl);
        const roleOk = allowed(a) && !(a.needsGraph && state.channel === "doc") && !state.pending && !a.comingSoon;
        const b = document.createElement("button");
        b.innerHTML = '<span class="lbl"></span><span class="hint"></span>';
        b.querySelector(".lbl").textContent = a.label;
        b.disabled = !(stageOk && roleOk);
        if (b.disabled) b.querySelector(".hint").textContent = a.comingSoon ? "Coming soon: auto-drafts the CCA-1 from this Appendix A and sends it through DocuSign"
          : !stageOk ? (WHEN[key] || "")
          : state.pending ? "Waiting for the last action to process"
          : (a.needsGraph && state.channel === "doc") ? "Needs the SharePoint connection (not set up yet)"
          : (step.who || "Not available to you");
        b.onclick = () => openAction(key, b);
        g.append(b);
      });
      box.append(g);
    });
    $("actionForm").hidden = true;
    show("actionSection", !!state.doc && state.doc.found !== false);
  }

  function openAction(key, btn) {
    const a = ACTIONS[key]; current = key;
    const docTest = state.channel === "doc" && /^26-205/.test(state.doc.subcontractNumber || "");
    const testTo = (state.settings && state.settings.TestExternalEmail) || "andrew.vanbeilen@gmail.com";
    const tradeTo = (state.test || docTest) ? testTo + " (TEST MODE)" : (state.doc.subEmail || "the subcontractor");
    $("actionPrompt").textContent = a.prompt + (key === "SendToTrade" ? tradeTo + "." : "");
    $("actionNote").value = ""; $("actionFile").value = "";
    $("actionNote").hidden = key === "SendToTrade" || key === "BuildPackage";
    $("fileLabel").hidden = !a.file || state.channel === "doc"; // attachments need the SharePoint connection
    $("numberLabel").hidden = !a.number;
    if (a.number) $("actionNumber").value = state.doc.subcontractNumber || "";
    $("codesLabel").hidden = !a.codes;
    if (a.codes) $("actionCodes").value = (state.record && state.record.CostCodes) || "";
    const sel = $("actionChoice"); sel.innerHTML = ""; sel.hidden = !a.choices;
    (a.choices || []).forEach((c) => { const o = document.createElement("option"); o.textContent = c; sel.append(o); });
    const form = $("actionForm");
    document.querySelectorAll("#actions button.chosen").forEach((x) => x.classList.remove("chosen"));
    if (btn) { btn.classList.add("chosen"); btn.after(form); }
    form.hidden = false;
    form.scrollIntoView({ block: "nearest", behavior: "smooth" });
    $("actionGo").focus({ preventScroll: true });
  }

  // Folder holding this Appendix A (e.g. Sub-Files/03200 Rebar (Shaw Steel)/Contract/Internal)
  const docFolder = (di) => decodeURIComponent(((di.parentReference && di.parentReference.path) || "").split("root:")[1] || "").replace(/^\/+/, "");

  async function uploadAttachment(file) {
    if (!file || state.channel === "doc") return "";
    const di = await DWGraph.driveItemFromUrl(state.docUrl);
    const folder = (state.project && state.project.CorrespondenceFolder) || docFolder(di);
    const bytes = new Uint8Array(await file.arrayBuffer());
    const name = `${state.record.Title} ${new Date().toISOString().slice(0, 10)} ${file.name}`;
    const up = await DWGraph.upload(di.parentReference.driveId, folder, name, bytes);
    return up.webUrl;
  }

  /** Replaces the subcontract number in every field of the Appendix A that holds it. */
  async function setDocumentNumber(n) {
    await Word.run(async (ctx) => {
      const ccs = ctx.document.contentControls.getByTag(C.tags.subcontractNumber);
      ccs.load("items");
      await ctx.sync();
      if (!ccs.items.length) throw new Error("Couldn't find the subcontract number field in this Appendix A.");
      ccs.items.forEach((cc) => cc.insertText(n, "Replace"));
      await ctx.sync();
    });
  }

  // Same status and event names the router uses, so the file can be updated the moment someone clicks.
  const STATUS_AFTER = {
    Submit: "Pending Approval", Resubmit: "Pending Approval", Approve: "Approved", ReturnToDraft: "Draft",
    SendToTrade: "With Trade", TradeAccepted: "With Accounting", TradeAlreadyAgreed: "With Accounting",
    TradeWantsChanges: "Draft", AccountingComplete: "Ready to Issue", ReturnFromAccounting: "Returned by Accounting", BackToAccounting: "With Accounting",
    Withdraw: "Draft", Cancel: "Cancelled", LogMarkup: "Signature Issue", LogDispute: "Signature Issue",
    "Resolve:Reissue unchanged": "Ready to Issue", "Resolve:Appendix A must change": "Draft", "Resolve:Cancel subcontract": "Cancelled"
  };
  const EVENT_NAME = {
    Submit: "Submitted", Resubmit: "Resubmitted", Approve: "Approved", ReturnToDraft: "Returned for changes",
    SendToTrade: "Sent to trade", TradeAccepted: "Trade accepted", TradeAlreadyAgreed: "Trade already agreed",
    TradeWantsChanges: "Trade wants changes", AccountingComplete: "Entered into accounting", ReturnFromAccounting: "Returned by accounting", BackToAccounting: "Back to accounting",
    Withdraw: "Withdrawn", Cancel: "Cancelled", LogMarkup: "Markup", LogDispute: "Dispute", Resolve: "Resolved", Comment: "Comment"
  };
  const pad = (n) => String(n).padStart(2, "0");
  function stamp(d) { d = d || new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`; }
  const oneLine = (t) => String(t || "").replace(/\s*\r?\n\s*/g, " ").replace(/\|/g, "/").trim();

  // The panel can't see who is signed in without the app registration, so it asks once and remembers.
  function storedName() { try { return localStorage.getItem("dwName") || ""; } catch (e) { return ""; } }
  function myName() {
    const n = ($("whoName").value || "").trim();
    if (!n) throw new Error("Enter your name at the top of the panel first (it goes in the history).");
    try { localStorage.setItem("dwName", n); } catch (e) { /* not kept; asked again next time */ }
    return n;
  }

  async function sendRequest(action, payload) {
    if (state.channel === "doc") {
      payload = payload || {};
      const before = state.record ? state.record.Status : "";
      const key = action === "Resolve" ? "Resolve:" + (payload.choice || "") : action;
      const who = myName();
      const note = oneLine(payload.note || payload.justification || "");
      const req = { id: "r" + Date.now() + Math.random().toString(36).slice(2, 6), action, subcontractNumber: state.doc.subcontractNumber,
        projectNumber: state.doc.projectNumber, documentUrl: state.docUrl, sentAt: new Date().toISOString(),
        payload: Object.assign({}, payload, { expectStatus: before, actorName: who }) };
      const line = [stamp(), who, EVENT_NAME[action] || action, note].join(" | ");
      const next = action === "Comment" ? null : (STATUS_AFTER[key] || before);
      const res = await DWDoc.apply(req, next, line);
      showDocState(res.status, res.history);
      return;
    }
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
    if (a.fileRequired && !file && state.channel !== "doc") return message("Please attach the email or file.", "error");
    if (state.preview) { $("actionForm").hidden = true; return message(`Preview: "${a.label}" would be recorded in the log and move this subcontract on. Nothing was sent.`, "info"); }
    $("actionGo").disabled = true;
    try {
      if (current === "BuildPackage") { await buildPackage(); }
      else {
        message("Working…", "info");
        let fullNote = note, newNumber;
        if (a.number) {
          const n = $("actionNumber").value.trim(), old = state.doc.subcontractNumber || "";
          if (!n) { $("actionGo").disabled = false; return message("Enter the subcontract number.", "error"); }
          if (n !== old) {
            await setDocumentNumber(n);
            newNumber = n;
            fullNote = `Subcontract number changed from ${old} to ${n}.` + (note ? " " + note : "");
          }
        }
        const link = await uploadAttachment(file);
        await sendRequest(current, { note: fullNote, attachmentUrl: link, choice: a.choices ? $("actionChoice").value : "",
          costCodes: a.codes ? $("actionCodes").value.trim() : undefined, newSubcontractNumber: newNumber });
        if (newNumber) { state.doc.subcontractNumber = newNumber; renderDoc(); }
        if (state.channel !== "doc") message(`${a.label} sent. The status will update in a minute.`, "ok");
      }
      $("actionForm").hidden = true;
      if (state.channel === "doc") docSent(a.label); else pollForChange();
    } catch (e) {
      message("Could not complete: " + e.message, "error");
    } finally { $("actionGo").disabled = false; }
  }

  async function backToAccounting() {
    if (state.preview) return message('Preview: "Back to Accounting" would send it straight back to accounting. Nothing was sent.', "info");
    $("backBtn").disabled = true;
    try {
      message("Working…", "info");
      await sendRequest("BackToAccounting", { note: $("backNote").value.trim() });
      $("backNote").value = "";
      if (state.channel === "doc") docSent("Back to Accounting");
      else { message("Back to Accounting sent. The status will update in a minute.", "ok"); pollForChange(); }
    } catch (e) {
      message("Could not complete: " + e.message, "error");
    } finally { $("backBtn").disabled = false; }
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
    const folder = state.project.PackagesFolder || docFolder(di);
    const up = await DWGraph.upload(di.parentReference.driveId, folder, DWPackage.packageFileName(sub), out.pdfBytes);
    await sendRequest("PackageBuilt", { packageUrl: up.webUrl, packageDriveId: di.parentReference.driveId, packageItemId: up.id, warnings: out.warnings });
    message("Package saved to " + folder + ". Open it from Status to review, then click Send for Signature." +
      (out.warnings.length ? " Check: " + out.warnings.join(" ") : ""), out.warnings.length ? "info" : "ok");
  }

  async function submit() {
    if (state.preview) return message("Preview only: open the Appendix A from SharePoint to submit.", "info");
    const d = state.doc;
    const est = parseMoney($("estimate").value);
    const codes = $("costCodes").value.trim();
    const just = $("justification").value.trim();
    if (isNaN(est)) return message("Enter the estimate carried.", "error");
    if (!codes) return message("Enter the cost code(s).", "error");
    const delta = Math.round((d.contractPrice - est) * 100) / 100;
    const gst = Math.round(d.contractPrice * 5) / 100;
    $("submitBtn").disabled = true;
    try {
      message("Submitting…", "info");
      await sendRequest(state.record ? "Resubmit" : "Submit", {
        estimateCarried: est, costCodes: codes, justification: just, note: just,
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
      if (state.channel === "doc") docSent("Submit"); else pollForChange();
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
      if (state.channel === "doc") docSent("Comment"); else setTimeout(refresh, 8000);
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
    const tb = $("testBanner");
    if (tb) { tb.hidden = !state.test; tb.textContent = state.test ? `TEST MODE: all emails and DocuSign go only to ${state.settings.TestInternalEmail} and ${state.settings.TestExternalEmail}.` : ""; }
    renderStatus();
    await loadLog();
  }

  // Preview only: pick a stage to see the buttons each person gets at that point.
  function setupStagePreview() {
    const box = document.createElement("section");
    box.innerHTML = '<h2>Preview a stage</h2><p class="muted">See the buttons at each step. Nothing is sent.</p>';
    const sel = document.createElement("select");
    ["Not submitted", "Pending Approval", "Approved", "With Trade", "With Accounting", "Returned by Accounting", "Ready to Issue",
      "Ready to Issue (package built)", "Out for Signature", "Signature Issue", "Executed"].forEach((t) => {
      const o = document.createElement("option"); o.textContent = t; sel.append(o);
    });
    sel.onchange = () => {
      const v = sel.value;
      message("", "info");
      if (v === "Not submitted") { state.record = null; renderStatus(); $("submitBtn").disabled = true; return; }
      const built = v.includes("package built");
      state.record = { Title: state.doc.subcontractNumber || "Preview", Status: built ? "Ready to Issue" : v,
        ContractValue: state.doc.contractPrice, PackageUrl: built ? "#" : "" };
      renderStatus();
      show("logSection", false);
      $("statusNote").textContent = "Preview of this stage. Approver and accounting buttons are shown together here; in real use each person only sees their own.";
    };
    box.append(sel);
    $("statusSection").before(box);
  }

  // What happens next at each status, shown under the status so everyone knows whose move it is.
  const NEXT = {
    "Draft": "Next: the PM revises the Appendix A and resubmits it for approval.",
    "Pending Approval": "Next: the approver clicks Approve in the Power Automate approval email (or here), or returns it to the PM here.",
    "Approved": "Next: the PM clicks Send to Trade, or Trade Already Agreed if the trade has already agreed to it.",
    "With Trade": "Next: when the trade replies, the PM clicks Trade Accepted or Trade Wants Changes.",
    "With Accounting": "Next: accounting enters the subcontract and clicks Entered into Accounting in the Power Automate approval email (or here), or sends it back to the PM here.",
    "Returned by Accounting": "Next: the PM fixes what accounting asked for and clicks Back to Accounting. No new approval needed unless the price or scope changed.",
    "Ready to Issue": "Next: the PM drafts the contract and sends it through DocuSign (this step is coming soon).",
    "Out for Signature": "Next: waiting on signatures through DocuSign.",
    "Signature Issue": "Next: the PM resolves the signature issue here.",
    "Executed": "Done: the subcontract is signed.",
    "Cancelled": "Nothing further: this subcontract is cancelled."
  };
  const NOT_SUBMITTED = "Not submitted yet. Submitting sends it to the approver, who gets an approval email with Approve buttons.";

  // ---------- document channel (no app registration) ----------
  function docSent(label) {
    message(`${label} done. It's logged and the emails go out within a minute or two. You can keep working.`, "ok");
  }

  function showDocState(status, history) {
    state.record = status ? { Title: state.doc.subcontractNumber, Status: status } : null;
    renderStatus();
    $("statusNote").textContent = status ? `${NEXT[status] || ""} Decisions made from an approval email show here once everyone has closed the file.`.trim() : NOT_SUBMITTED;
    const list = $("log"); list.innerHTML = "";
    (history || "").split(/\r?\n/).filter((l) => l.trim()).reverse().forEach((line) => {
      const [when, who, what, note] = line.split(" | ");
      const li = document.createElement("li");
      li.innerHTML = `<div class="what"></div><div class="when"></div><div class="note"></div>`;
      li.querySelector(".what").textContent = what || line;
      li.querySelector(".when").textContent = [when, who].filter(Boolean).join(" · ");
      li.querySelector(".note").textContent = note || "";
      list.append(li);
    });
    show("logSection", list.children.length > 0);
  }

  async function loadFromDoc(d) {
    state.channel = "doc";
    state.pending = false;
    const saved = storedName();
    $("whoName").value = saved || await DWDoc.lastAuthor();
    show("whoRow", true);
    showDocState(d.status, d.history);
    const tb = $("testBanner");
    if (tb && /^26-205/.test(state.doc.subcontractNumber || "")) { tb.hidden = false; tb.textContent = "TEST MODE: emails go only to Andrew's work and Gmail addresses."; }
  }

  async function start() {
    try {
      state.docUrl = Office.context.document.url || "";
      state.doc = await readDocument();
      show("loading", false);
      renderDoc();
      const notConnected = /^0{8}-/.test(C.clientId || "0");
      if (/^https:\/\//i.test(state.docUrl) && notConnected) {
        let d = null, why = "";
        try { d = await DWDoc.read(8); } catch (e) { d = null; why = e.message || String(e); }
        if (d) { await loadFromDoc(d); return; }
        show("loading", false);
        message("Couldn't read the approval fields from this file" + (why ? " (" + why + ")" : "") +
          ". Close the file, wait a few seconds, then open it again from SharePoint. If this keeps happening, tell Andrew.", "error");
        return;
      }
      if (!/^https:\/\//i.test(state.docUrl) || notConnected) {
        // Preview mode: show the full submit form so it can be tried, but block sending.
        state.preview = true;
        message(notConnected
          ? "Preview only: the panel isn't connected to SharePoint yet (waiting on the app registration). Reading the Appendix A works; use Preview a stage below to click through."
          : "Preview only: this copy isn't saved in SharePoint, so it can't be submitted. Open the Appendix A from a project folder to submit.", "info");
        state.record = null;
        renderStatus();
        $("statusNote").textContent = notConnected ? "Preview. Submitting needs the SharePoint connection." : "Preview. Submitting needs the file in SharePoint.";
        const b = $("submitBtn");
        b.disabled = true;
        b.textContent = "Submit for Approval (needs SharePoint)";
        setupStagePreview();
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
    $("backBtn").onclick = backToAccounting;
    $("actionCancel").onclick = () => { $("actionForm").hidden = true; };
    $("commentBtn").onclick = addComment;
    start();
  });
})();
