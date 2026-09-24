/*
 * Sign-in and Microsoft Graph helpers for the add-in.
 * Uses Nested App Authentication (the user's existing Office sign-in), so there is no separate login.
 */
(function () {
  const C = window.DW_CONFIG;
  const G = "https://graph.microsoft.com/v1.0";
  let pca = null;
  let siteId = null;
  const listIds = {};

  async function initAuth() {
    if (pca) return pca;
    pca = await msal.createNestablePublicClientApplication({
      auth: { clientId: C.clientId, authority: "https://login.microsoftonline.com/organizations" }
    });
    return pca;
  }

  async function token() {
    await initAuth();
    const req = { scopes: C.graphScopes };
    const acct = pca.getActiveAccount() || pca.getAllAccounts()[0];
    try {
      const r = await pca.acquireTokenSilent(Object.assign({ account: acct }, req));
      return r.accessToken;
    } catch (e) {
      const r = await pca.acquireTokenPopup(req);
      pca.setActiveAccount(r.account);
      return r.accessToken;
    }
  }

  async function call(method, url, body, opts) {
    opts = opts || {};
    const headers = Object.assign({ Authorization: "Bearer " + (await token()) }, opts.headers || {});
    let payload = body;
    if (body && !(body instanceof ArrayBuffer) && !(body instanceof Uint8Array) && !(body instanceof Blob)) {
      headers["Content-Type"] = "application/json";
      payload = JSON.stringify(body);
    }
    const res = await fetch(url.startsWith("http") ? url : G + url, { method, headers, body: payload });
    if (!res.ok) {
      let msg = res.status + " " + res.statusText;
      try { const j = await res.json(); if (j.error && j.error.message) msg = j.error.message; } catch (e) { /* not JSON */ }
      throw new Error(msg);
    }
    if (opts.raw) return res;
    if (res.status === 204) return null;
    const ct = res.headers.get("content-type") || "";
    return ct.includes("json") ? res.json() : res.arrayBuffer();
  }

  async function hubSiteId() {
    if (siteId) return siteId;
    const s = await call("GET", `/sites/${C.sharepointHost}:${C.hubSitePath}`);
    siteId = s.id;
    return siteId;
  }

  async function listId(key) {
    if (listIds[key]) return listIds[key];
    const sid = await hubSiteId();
    const name = C.lists[key];
    const r = await call("GET", `/sites/${sid}/lists?$filter=displayName eq '${name.replace(/'/g, "''")}'&$select=id,displayName`);
    if (!r.value.length) throw new Error(`List "${name}" not found on the hub site.`);
    listIds[key] = r.value[0].id;
    return listIds[key];
  }

  const esc = (v) => String(v).replace(/'/g, "''");

  /** Items from a hub list. filter uses Graph syntax on fields, e.g. "fields/Title eq 'X'". */
  async function items(key, filter, orderby, top) {
    const sid = await hubSiteId();
    const lid = await listId(key);
    let url = `/sites/${sid}/lists/${lid}/items?$expand=fields&$top=${top || 200}`;
    if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
    if (orderby) url += `&$orderby=${encodeURIComponent(orderby)}`;
    const r = await call("GET", url, null, { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } });
    return r.value.map((i) => Object.assign({ _id: i.id, _createdBy: i.createdBy && i.createdBy.user }, i.fields));
  }

  async function createItem(key, fields) {
    const sid = await hubSiteId();
    const lid = await listId(key);
    return call("POST", `/sites/${sid}/lists/${lid}/items`, { fields });
  }

  async function me() { return call("GET", "/me?$select=displayName,mail,userPrincipalName"); }

  /** Resolves a SharePoint file URL (as Word reports it) to its drive item. */
  async function driveItemFromUrl(fileUrl) {
    const clean = fileUrl.split("?")[0];
    const b64 = btoa(unescape(encodeURIComponent(clean))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
    return call("GET", `/shares/u!${b64}/driveItem?$select=id,name,webUrl,parentReference,file`);
  }

  /** File contents as PDF (Word files are converted by SharePoint). */
  async function fileAsPdf(driveId, itemId) {
    return call("GET", `/drives/${driveId}/items/${itemId}/content?format=pdf`, null, {}).then((b) => new Uint8Array(b));
  }

  async function fileBytes(driveId, itemId) {
    return call("GET", `/drives/${driveId}/items/${itemId}/content`).then((b) => new Uint8Array(b));
  }

  /** Uploads bytes to a folder in a drive; uses an upload session for files over 4 MB. */
  async function upload(driveId, folderPath, fileName, bytes) {
    const path = [folderPath.replace(/^\/+|\/+$/g, ""), fileName].filter(Boolean).map(encodeURIComponent).join("/");
    if (bytes.byteLength < 4 * 1024 * 1024) {
      return call("PUT", `/drives/${driveId}/root:/${path}:/content?@microsoft.graph.conflictBehavior=rename`, bytes);
    }
    const s = await call("POST", `/drives/${driveId}/root:/${path}:/createUploadSession`,
      { item: { "@microsoft.graph.conflictBehavior": "rename" } });
    const chunk = 5 * 327680;
    let result = null;
    for (let start = 0; start < bytes.byteLength; start += chunk) {
      const end = Math.min(start + chunk, bytes.byteLength);
      const res = await fetch(s.uploadUrl, {
        method: "PUT",
        headers: { "Content-Range": `bytes ${start}-${end - 1}/${bytes.byteLength}` },
        body: bytes.slice(start, end)
      });
      if (!res.ok) throw new Error("Upload failed: " + res.status);
      if (res.status === 200 || res.status === 201) result = await res.json();
    }
    return result;
  }

  window.DWGraph = { token, call, items, createItem, me, driveItemFromUrl, fileAsPdf, fileBytes, upload, esc, hubSiteId };
})();
