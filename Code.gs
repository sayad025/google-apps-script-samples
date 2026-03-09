/****************************************************
 * SAVFLOW PRO 3 — Code.gs
 * Google Apps Script Backend
 * Sheets: USERS, PROJECTS, APPARTEMENTS, CORPS_ETAT,
 *         DESIGNATIONS, LOCALISATIONS, PRESTATAIRES,
 *         RECEPTIONS, RECEPTION_ITEMS, DEBUG_LOGS
 ****************************************************/

/* ============================================================
   SHEET NAMES — must match exactly
============================================================ */
const S = {
  USERS:      "USERS",
  PROJECTS:   "PROJECTS",
  APPTS:      "APPARTEMENTS",
  CORPS:      "CORPS_ETAT",
  DESIG:      "DESIGNATIONS",
  LOCS:       "LOCALISATIONS",
  PREST:      "PRESTATAIRES",
  RECEP:      "RECEPTIONS",
  ITEMS:      "RECEPTION_ITEMS",
  DBG:        "DEBUG_LOGS"
};

const DRIVE_ROOT = "SAVFLOW_FILES";

/* ============================================================
   ENTRY POINTS
============================================================ */
function doGet()  { return _j({ ok: true, service: "SAVFLOW_PRO3", v: "3.0" }); }

function doPost(e) {
  const t0 = Date.now();
  try {
    const req    = _parse(e);
    const action = _s(req.action);
    _dbg("CALL", { action, email: _s(req.email) });

    if (!action) return _j({ ok: false, error: "Missing action" });

    let r;
    switch (action) {
      /* Auth */
      case "login":                r = _login(req);              break;
      /* Master data */
      case "get_projects":         r = _getProjects(req);        break;
      case "get_appartements":     r = _getAppartements(req);    break;
      case "get_corps_etat":       r = _getCorps(req);           break;
      case "get_designations":     r = _getDesig(req);           break;
      case "get_localisations":    r = _getLocs(req);            break;
      case "get_prestataires":     r = _getPrestataires(req);    break;
      /* Tickets */
      case "create_ticket":        r = _createTicket(req);       break;
      case "get_tickets":          r = _getTickets(req);         break;
      case "update_ticket":        r = _updateTicket(req);       break;
      /* Photos */
      case "upload_photos":        r = _uploadPhotos(req);       break;
      /* PDF */
      case "generate_pdf":         r = _generatePdf(req);        break;
      /* Dashboard */
      case "get_dashboard":        r = _getDashboard(req);       break;

      default: r = { ok: false, error: "Unknown action: " + action };
    }

    _dbg("DONE", { action, ms: Date.now() - t0, ok: Boolean(r && r.ok) });
    return _j(r);
  } catch(err) {
    const msg = err && err.message ? err.message : String(err);
    _dbg("ERR", { message: msg, stack: err && err.stack ? String(err.stack).slice(0,800) : "" });
    return _j({ ok: false, error: msg });
  }
}

/* ============================================================
   PARSE / RESPOND
============================================================ */
function _parse(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
  let p = {};
  try { p = JSON.parse(raw); } catch(_) {}
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(k => { if (p[k] === undefined) p[k] = e.parameter[k]; });
  }
  return (p && typeof p === "object") ? p : {};
}

function _j(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   SHEET HELPERS
============================================================ */
function _ss()             { return SpreadsheetApp.getActiveSpreadsheet(); }
function _sh(name, create) {
  const ss = _ss(); let sh = ss.getSheetByName(name);
  if (!sh && create) sh = ss.insertSheet(name);
  return sh;
}

function _all(sh) {
  if (!sh) return { headers: [], rows: [] };
  const last = sh.getLastRow();
  if (last < 1) return { headers: [], rows: [] };
  const data    = sh.getDataRange().getValues();
  const headers = data[0].map(h => _s(h));
  const rows    = last > 1 ? data.slice(1) : [];
  return { headers, rows };
}

function _hmap(headers) {
  const m = {};
  headers.forEach((h, i) => { if (h) m[h] = i; });
  return m;
}

function _append(sh, obj) {
  const { headers } = _all(sh);
  const m = _hmap(headers);
  const row = new Array(headers.length).fill("");
  Object.keys(obj).forEach(k => {
    if (m[k] !== undefined) row[m[k]] = (obj[k] === null || obj[k] === undefined) ? "" : obj[k];
  });
  sh.appendRow(row);
}

function _rowObj(headers, row) {
  const m = _hmap(headers);
  const o = {};
  headers.forEach((h, i) => { o[h] = row[i]; });
  return o;
}

/* ============================================================
   LOGGER
============================================================ */
function _dbg(tag, obj) {
  try {
    let sh = _sh(S.DBG, true);
    if (sh.getLastRow() === 0) sh.appendRow(["Timestamp","Tag","Data"]);
    sh.appendRow([new Date(), _s(tag), JSON.stringify(obj || {})]);
    Logger.log("[%s] %s", tag, JSON.stringify(obj || {}));
  } catch(_) {}
}

/* ============================================================
   STRING / DATE UTILS
============================================================ */
function _s(v)   { return String(v == null ? "" : v).trim(); }
function _su(v)  { return _s(v).toUpperCase(); }
function _bool(v){ return _s(v).toUpperCase() === "TRUE"; }
function _ts()   {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}
function _fmt(d, fmt) {
  if (!d) return "";
  try { return Utilities.formatDate(d instanceof Date ? d : new Date(d), Session.getScriptTimeZone(), fmt || "dd/MM/yyyy HH:mm"); }
  catch(_) { return _s(d); }
}
function _safe(s) { return _s(s).replace(/[\/\\?%*:|"<>\n\r\t]/g,"_").slice(0,100) || "file"; }

function _ticketId(projectId) {
  const d   = new Date();
  const yy  = String(d.getFullYear()).slice(-2);
  const mm  = ("0"+(d.getMonth()+1)).slice(-2);
  const dd  = ("0"+d.getDate()).slice(-2);
  const ms  = String(Date.now()).slice(-4);
  return (projectId || "SAV").slice(0,6).toUpperCase() + "-" + yy + mm + dd + "-" + ms;
}

function _diffMin(a, b) {
  try {
    const da = a instanceof Date ? a : new Date(a);
    const db = b instanceof Date ? b : new Date(b);
    if (isNaN(da) || isNaN(db)) return null;
    return Math.round((db - da) / 60000);
  } catch(_) { return null; }
}

/* ============================================================
   AUTH
============================================================ */
function _login(req) {
  const email = _s(req.email).toLowerCase();
  if (!email) return { ok: false, error: "Email requis" };

  const sh = _sh(S.USERS);
  if (!sh) return { ok: false, error: "Sheet USERS introuvable" };

  const { headers, rows } = _all(sh);
  const m = _hmap(headers);

  const iE = m["Email"]; const iR = m["Role"]; const iA = m["IsActive"];
  if (iE === undefined || iR === undefined || iA === undefined)
    return { ok: false, error: "USERS: headers Email|Role|IsActive requis" };

  for (const row of rows) {
    if (_s(row[iE]).toLowerCase() !== email) continue;
    if (!_bool(row[iA])) return { ok: false, error: "Compte inactif" };
    return { ok: true, role: _su(row[iR]) || "AGENT" };
  }
  return { ok: false, error: "Email non autorisé" };
}

/* ============================================================
   MASTER DATA
============================================================ */
function _getProjects(req) {
  const sh = _sh(S.PROJECTS);
  if (!sh) return { ok: false, error: "Sheet PROJECTS introuvable" };
  const { headers, rows } = _all(sh);
  const m = _hmap(headers);
  const items = rows
    .filter(r => m["IsActive"] === undefined || _bool(r[m["IsActive"]]))
    .map(r => ({ id: _s(r[m["ProjectID"]]), name: _s(r[m["ProjectName"]]) }))
    .filter(p => p.id);
  return { ok: true, items };
}

function _getAppartements(req) {
  const sh = _sh(S.APPTS);
  if (!sh) return { ok: false, error: "Sheet APPARTEMENTS introuvable" };
  const proj = _su(req.projectId || "");
  const { headers, rows } = _all(sh);
  const m = _hmap(headers);
  const iK = m["AppartementKey"]; const iP = m["ProjectID"]; const iA = m["IsActive"];
  if (iK === undefined) return { ok: false, error: "Header AppartementKey manquant" };
  const items = rows
    .filter(r => (iA === undefined || _bool(r[iA])))
    .filter(r => !proj || proj === "ALL" || _su(r[iP]) === proj || _su(r[iP]) === "ALL")
    .map(r => _s(r[iK]))
    .filter(Boolean);
  return { ok: true, items: _uniq(items) };
}

function _getCorps(req) {
  const sh = _sh(S.CORPS);
  if (!sh) return { ok: false, error: "Sheet CORPS_ETAT introuvable" };
  const proj = _su(req.projectId || "");
  const { headers, rows } = _all(sh);
  const m = _hmap(headers);
  const iC = m["CorpsEtat"]; const iP = m["ProjectID"]; const iA = m["IsActive"];
  if (iC === undefined) return { ok: false, error: "Header CorpsEtat manquant" };
  const items = rows
    .filter(r => (iA === undefined || _bool(r[iA])))
    .filter(r => !proj || proj === "ALL" || _su(r[iP]) === proj || _su(r[iP]) === "ALL")
    .map(r => _s(r[iC]))
    .filter(Boolean);
  _dbg("CORPS", { proj, count: items.length });
  return { ok: true, items: _uniq(items) };
}

function _getDesig(req) {
  const sh = _sh(S.DESIG);
  if (!sh) return { ok: false, error: "Sheet DESIGNATIONS introuvable" };
  const proj     = _su(req.projectId || "");
  const corpsEtat = _su(req.corpsEtat || "");
  const { headers, rows } = _all(sh);
  const m = _hmap(headers);
  const iP = m["ProjectID"]; const iC = m["CorpsEtat"]; const iD = m["Designation"];
  const iA = m["IsActive"];  const iO = m["Ordre"];
  if (iC === undefined || iD === undefined) return { ok: false, error: "Headers DESIGNATIONS: CorpsEtat|Designation requis" };

  let filtered = rows.filter(r => (iA === undefined || _bool(r[iA])));
  if (proj  && proj  !== "ALL") filtered = filtered.filter(r => _su(r[iP]) === proj || _su(r[iP]) === "ALL");
  if (corpsEtat)                filtered = filtered.filter(r => _su(r[iC]) === corpsEtat);

  if (iO !== undefined) filtered.sort((a,b) => Number(a[iO]||999) - Number(b[iO]||999));

  const items = _uniq(filtered.map(r => _s(r[iD])).filter(Boolean));
  _dbg("DESIG", { corps: corpsEtat, proj, count: items.length, first3: items.slice(0,3) });
  return { ok: true, items };
}

function _getLocs(req) {
  const sh = _sh(S.LOCS);
  if (!sh) return { ok: false, error: "Sheet LOCALISATIONS introuvable" };
  const proj = _su(req.projectId || "");
  const { headers, rows } = _all(sh);
  const m = _hmap(headers);
  const iP = m["ProjectID"]; const iL = m["Localisation"]; const iA = m["IsActive"]; const iO = m["Ordre"];
  if (iL === undefined) return { ok: false, error: "Header Localisation manquant" };

  let filtered = rows.filter(r => (iA === undefined || _bool(r[iA])));
  if (proj && proj !== "ALL") filtered = filtered.filter(r => _su(r[iP]) === proj || _su(r[iP]) === "ALL");
  if (iO !== undefined) filtered.sort((a,b) => Number(a[iO]||999) - Number(b[iO]||999));

  const items = _uniq(filtered.map(r => _s(r[iL])).filter(Boolean));
  _dbg("LOCS", { proj, count: items.length, first5: items.slice(0,5) });
  return { ok: true, items };
}

function _getPrestataires(req) {
  const sh = _sh(S.PREST);
  if (!sh) return { ok: false, error: "Sheet PRESTATAIRES introuvable" };
  const corps = _su(req.corpsEtat || "");
  const { headers, rows } = _all(sh);
  const m = _hmap(headers);
  const iN = m["Name"]; const iC = m["CorpsEtat"]; const iA = m["IsActive"];
  if (iN === undefined) return { ok: false, error: "Header Name manquant" };

  let filtered = rows.filter(r => (iA === undefined || _bool(r[iA])));
  if (corps) filtered = filtered.filter(r => !_s(r[iC]) || _su(r[iC]) === corps || _su(r[iC]) === "ALL");

  const items = _uniq(filtered.map(r => _s(r[iN])).filter(Boolean));
  return { ok: true, items };
}

/* ============================================================
   DRIVE
============================================================ */
function _rootFolder() {
  const it = DriveApp.getFoldersByName(DRIVE_ROOT);
  if (it.hasNext()) return it.next();
  const f = DriveApp.createFolder(DRIVE_ROOT);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f;
}

function _subFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  const f = parent.createFolder(name);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f;
}

function _ticketFolder(projectId, ticketId) {
  const root = _rootFolder();
  const fp   = _subFolder(root, projectId || "DEFAULT");
  return _subFolder(fp, ticketId);
}

/* ============================================================
   UPLOAD PHOTOS
============================================================ */
function _uploadPhotos(req) {
  const projectId = _s(req.projectId) || "DEFAULT";
  const ticketId  = _s(req.ticketId)  || _ticketId(projectId);
  const files     = Array.isArray(req.files) ? req.files : [];

  if (!files.length) return { ok: false, error: "Aucun fichier" };
  if (files.length > 8) return { ok: false, error: "Max 8 photos par upload" };

  const folder = _ticketFolder(projectId, ticketId);
  const out = [];

  files.forEach((f, i) => {
    const name  = _safe(f.name || ("photo_"+(i+1)+".jpg"));
    const mime  = _s(f.mimeType) || "image/jpeg";
    const b64   = _s(f.dataBase64);
    if (!b64) return;
    try {
      const bytes = Utilities.base64Decode(b64);
      const blob  = Utilities.newBlob(bytes, mime, name);
      const file  = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      out.push({ name: file.getName(), url: file.getUrl() });
    } catch(e) {
      _dbg("UPLOAD_ERR", { file: name, error: e.message });
    }
  });

  _dbg("PHOTOS", { ticketId, uploaded: out.length, total: files.length });

  return {
    ok: true,
    ticketId,
    folderUrl: folder.getUrl(),
    files:  out,
    titles: out.map(x => x.name),
    links:  out.map(x => x.url)
  };
}

/* ============================================================
   CREATE TICKET
   Writes to RECEPTIONS + RECEPTION_ITEMS
============================================================ */
function _createTicket(req) {
  const shR = _sh(S.RECEP);
  const shI = _sh(S.ITEMS);
  if (!shR) return { ok: false, error: "Sheet RECEPTIONS introuvable" };
  if (!shI) return { ok: false, error: "Sheet RECEPTION_ITEMS introuvable" };

  const type   = _su(req.type || "SAV");    // SAV | LIVRAISON | CHANTIER
  const projId = _s(req.projectId) || "DEFAULT";
  const appt   = _s(req.appartementKey);
  if (!appt) return { ok: false, error: "AppartementKey requis" };

  const items = Array.isArray(req.items) ? req.items : [];
  if (!items.length) return { ok: false, error: "Au moins 1 item requis (corps d'état)" };

  const ticketId = _ticketId(projId);
  const now      = _ts();

  _append(shR, {
    TicketID:        ticketId,
    Type:            type,
    ProjectID:       projId,
    AppartementKey:  appt,
    ClientName:      _s(req.clientName),
    CreatedAt:       now,
    CreatedBy:       _s(req.email),
    Collaborateur:   _s(req.email),
    Prestataire:     _s(req.prestataire),
    Statut:          "Nouvelle",
    Etape:           "Diagnostic",
    PlannedAt:       "",
    StartAt:         "",
    DoneAt:          "",
    CommentaireTech: _s(req.commentaireTech),
    PhotoLinks:      Array.isArray(req.photoLinks) ? req.photoLinks.join("\n") : _s(req.photoLinks),
    PdfLink:         ""
  });

  items.forEach(it => {
    _append(shI, {
      TicketID:      ticketId,
      CorpsEtat:     _s(it.corpsEtat),
      Designation:   _s(it.designation),
      Localisation:  _s(it.localisation),
      Remarque:      _s(it.remarque),
      IsFixRequired: it.isFixRequired ? "TRUE" : "FALSE"
    });
  });

  _dbg("CREATE", { ticketId, type, appt, items: items.length });
  return { ok: true, ticketId };
}

/* ============================================================
   GET TICKETS  (Suivi + joins RECEPTION_ITEMS)
============================================================ */
function _getTickets(req) {
  const shR = _sh(S.RECEP);
  const shI = _sh(S.ITEMS);
  if (!shR) return { ok: false, error: "Sheet RECEPTIONS introuvable" };

  const filterProj  = _su(req.projectId || "");
  const filterType  = _su(req.type || "");
  const filterStat  = _s(req.statut || "");
  const filterAppt  = _s(req.appartementKey || "");
  const limit       = Math.min(500, Math.max(1, Number(req.limit || 200)));

  const { headers: rH, rows: rRows } = _all(shR);
  const rM = _hmap(rH);

  // Build items map
  const itemsMap = {};
  if (shI) {
    const { headers: iH, rows: iRows } = _all(shI);
    const iM = _hmap(iH);
    iRows.forEach(r => {
      const tid = _s(r[iM["TicketID"]]);
      if (!tid) return;
      if (!itemsMap[tid]) itemsMap[tid] = [];
      itemsMap[tid].push({
        corpsEtat:    _s(r[iM["CorpsEtat"]]),
        designation:  _s(r[iM["Designation"]]),
        localisation: _s(r[iM["Localisation"]]),
        remarque:     _s(r[iM["Remarque"]]),
        isFixRequired: _bool(r[iM["IsFixRequired"]])
      });
    });
  }

  const out = [];
  for (let i = rRows.length - 1; i >= 0 && out.length < limit; i--) {
    const row = rRows[i];
    if (filterProj && filterProj !== "ALL" && _su(row[rM["ProjectID"]]) !== filterProj) continue;
    if (filterType && _su(row[rM["Type"]]) !== filterType) continue;
    if (filterStat && _s(row[rM["Statut"]]) !== filterStat) continue;
    if (filterAppt && _s(row[rM["AppartementKey"]]) !== filterAppt) continue;

    const tid      = _s(row[rM["TicketID"]]);
    const createdAt = row[rM["CreatedAt"]];
    const plannedAt = row[rM["PlannedAt"]];
    const startAt   = row[rM["StartAt"]];
    const doneAt    = row[rM["DoneAt"]];

    const totalMin  = _diffMin(createdAt, doneAt);
    const repairMin = _diffMin(startAt, doneAt);
    const isLate    = (plannedAt && doneAt) ? (new Date(doneAt) > new Date(plannedAt)) : false;

    out.push({
      ticketId:       tid,
      type:           _s(row[rM["Type"]]),
      projectId:      _s(row[rM["ProjectID"]]),
      appartementKey: _s(row[rM["AppartementKey"]]),
      clientName:     _s(row[rM["ClientName"]]),
      createdAt:      _fmt(createdAt),
      createdBy:      _s(row[rM["CreatedBy"]]),
      collaborateur:  _s(row[rM["Collaborateur"]]),
      prestataire:    _s(row[rM["Prestataire"]]),
      statut:         _s(row[rM["Statut"]]),
      etape:          _s(row[rM["Etape"]]),
      plannedAt:      _fmt(plannedAt),
      startAt:        _fmt(startAt),
      doneAt:         _fmt(doneAt),
      commentaireTech: _s(row[rM["CommentaireTech"]]),
      photoLinks:     _s(row[rM["PhotoLinks"]]),
      pdfLink:        _s(row[rM["PdfLink"]]),
      totalMin,
      repairMin,
      isLate,
      items:          itemsMap[tid] || []
    });
  }

  return { ok: true, items: out, total: out.length };
}

/* ============================================================
   UPDATE TICKET  (statut, etape, plannedAt, comment, etc.)
============================================================ */
function _updateTicket(req) {
  const shR = _sh(S.RECEP);
  if (!shR) return { ok: false, error: "Sheet RECEPTIONS introuvable" };

  const ticketId = _s(req.ticketId);
  if (!ticketId) return { ok: false, error: "ticketId requis" };

  const { headers, rows } = _all(shR);
  const m  = _hmap(headers);
  const iT = m["TicketID"];
  if (iT === undefined) return { ok: false, error: "Header TicketID manquant" };

  const fields = ["Statut","Etape","PlannedAt","StartAt","DoneAt","CommentaireTech","Prestataire","PhotoLinks","PdfLink"];

  // Auto-fill timestamps
  const statut = _s(req.statut);
  const now    = _ts();
  if (statut === "En cours"  && !req.startAt) req.startAt = now;
  if (statut === "Terminée"  && !req.doneAt)  req.doneAt  = now;

  for (let i = 0; i < rows.length; i++) {
    if (_s(rows[i][iT]) !== ticketId) continue;
    const rowNum = i + 2; // 1-indexed, skip header
    fields.forEach(f => {
      if (req[_lcFirst(f)] !== undefined && m[f] !== undefined) {
        shR.getRange(rowNum, m[f] + 1).setValue(_s(req[_lcFirst(f)]));
      }
    });
    _dbg("UPDATE", { ticketId, fields: Object.keys(req).filter(k => fields.includes(_ucFirst(k))) });
    return { ok: true };
  }
  return { ok: false, error: "Ticket introuvable: " + ticketId };
}

function _lcFirst(s) { return s.charAt(0).toLowerCase() + s.slice(1); }
function _ucFirst(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

/* ============================================================
   PDF GENERATION
============================================================ */
function _generatePdf(req) {
  const ticketId  = _s(req.ticketId);
  const type      = _su(req.type || "SAV");
  const projId    = _s(req.projectId) || "DEFAULT";
  const appt      = _s(req.appartementKey);
  const client    = _s(req.clientName);
  const prest     = _s(req.prestataire);
  const email     = _s(req.email);
  const items     = Array.isArray(req.items) ? req.items : [];
  const photoLinks = _s(req.photoLinks);
  const commentaire = _s(req.commentaireTech);

  if (!appt) return { ok: false, error: "AppartementKey requis" };

  const now     = new Date();
  const dateStr = _fmt(now, "dd/MM/yyyy HH:mm");
  const folder  = _ticketFolder(projId, ticketId || _ticketId(projId));
  const pdfName = _safe(`PV_${type}_${appt}_${Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmm")}.pdf`);

  const html = _buildPdfHtml({ type, projId, appt, client, prest, dateStr, ticketId, items, photoLinks, commentaire, email });
  const blob  = HtmlService.createHtmlOutput(html).getBlob().setName(pdfName).getAs(MimeType.PDF);
  const file  = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const pdfUrl = file.getUrl();

  // Update RECEPTIONS.PdfLink if ticketId exists
  if (ticketId) {
    const shR = _sh(S.RECEP);
    if (shR) {
      const { headers, rows } = _all(shR);
      const m = _hmap(headers); const iT = m["TicketID"]; const iP = m["PdfLink"];
      if (iT !== undefined && iP !== undefined) {
        for (let i = 0; i < rows.length; i++) {
          if (_s(rows[i][iT]) === ticketId) {
            shR.getRange(i + 2, iP + 1).setValue(pdfUrl);
            break;
          }
        }
      }
    }
  }

  return { ok: true, pdfUrl, fileName: file.getName() };
}

/* ============================================================
   PDF HTML BUILDER
============================================================ */
const PDF_CSS = `
body{font-family:Arial,sans-serif;padding:28px;color:#1a1a1a;margin:0;font-size:13px}
.hdr{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #333;padding-bottom:14px;margin-bottom:18px}
.hdr-left h1{margin:0 0 4px;font-size:20px}
.hdr-left p{margin:2px 0;color:#555;font-size:12px}
.badge{display:inline-block;padding:5px 12px;border-radius:6px;font-weight:800;font-size:13px;letter-spacing:.5px}
.badge-sav{background:#fef3cd;border:1px solid #e6b800;color:#7a5900}
.badge-liv{background:#d4edda;border:1px solid #28a745;color:#145523}
.badge-chan{background:#d1ecf1;border:1px solid #0288d1;color:#014361}
.meta{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:18px}
.meta-item{padding:8px 12px;background:#f7f7f7;border-radius:8px;font-size:12px}
.meta-item b{display:block;color:#666;font-size:10px;text-transform:uppercase;margin-bottom:2px}
table{width:100%;border-collapse:collapse;margin-bottom:16px;font-size:12px}
th{background:#333;color:#fff;padding:8px;text-align:left;font-weight:700}
td{padding:8px;border-bottom:1px solid #eee}
tr:nth-child(even){background:#f9f9f9}
.section-title{font-weight:800;font-size:13px;text-transform:uppercase;color:#333;border-bottom:1px solid #ddd;padding-bottom:4px;margin:16px 0 8px}
.sig-area{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-top:24px}
.sig-box{border:1px solid #ccc;border-radius:8px;padding:12px;min-height:80px}
.sig-label{font-size:11px;color:#888;margin-bottom:6px;text-transform:uppercase}
.footer{margin-top:24px;text-align:center;font-size:10px;color:#aaa;border-top:1px solid #eee;padding-top:10px}
a{color:#0b6;text-decoration:none}
`;

function _e(s)  { return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
function _ea(s) { return String(s||"").replace(/"/g,"&quot;"); }

function _buildPdfHtml(d) {
  const badgeCls = d.type==="LIVRAISON" ? "badge-liv" : d.type==="CHANTIER" ? "badge-chan" : "badge-sav";
  const typeLabel = d.type==="LIVRAISON" ? "Réception Client" : d.type==="CHANTIER" ? "Réception Chantier" : "SAV Réclamation";

  const itemRows = (d.items||[]).map(it =>
    `<tr><td>${_e(it.corpsEtat)}</td><td>${_e(it.designation)}</td><td>${_e(it.localisation)}</td><td>${_e(it.remarque)}</td></tr>`
  ).join("") || `<tr><td colspan="4" style="color:#999;text-align:center">Aucun item</td></tr>`;

  const photoHtml = (d.photoLinks||"").split("\n").map(u=>u.trim()).filter(Boolean)
    .map((u,i)=>`<div><a href="${_ea(u)}" target="_blank">Photo ${i+1}: ${_e(u)}</a></div>`).join("") ||
    `<span style="color:#999">Aucune photo.</span>`;

  const sigBoxes = d.type === "SAV"
    ? `<div class="sig-box"><div class="sig-label">Technicien</div></div>
       <div class="sig-box"><div class="sig-label">Responsable SAV</div></div>
       <div class="sig-box"><div class="sig-label">Client (si présent)</div></div>`
    : d.type === "LIVRAISON"
    ? `<div class="sig-box"><div class="sig-label">Client</div></div>
       <div class="sig-box"><div class="sig-label">Responsable Commercial</div></div>
       <div class="sig-box"><div class="sig-label">Responsable Technique</div></div>`
    : `<div class="sig-box"><div class="sig-label">Prestataire / Cachet</div></div>
       <div class="sig-box"><div class="sig-label">Chef de Chantier</div></div>
       <div class="sig-box"><div class="sig-label">Responsable Qualité</div></div>`;

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>PV ${_e(typeLabel)}</title>
<style>${PDF_CSS}</style></head><body>
<div class="hdr">
  <div class="hdr-left">
    <h1>PV ${_e(typeLabel)}</h1>
    <p>Projet: <b>${_e(d.projId)}</b> &nbsp;|&nbsp; Appartement: <b>${_e(d.appt)}</b></p>
    <p>Ticket: <b>${_e(d.ticketId||"-")}</b> &nbsp;|&nbsp; Date: ${_e(d.dateStr)}</p>
    ${d.client ? `<p>Client: <b>${_e(d.client)}</b></p>` : ""}
    ${d.prest  ? `<p>Prestataire: <b>${_e(d.prest)}</b></p>` : ""}
    <p>Généré par: ${_e(d.email||"-")}</p>
  </div>
  <div><span class="badge ${badgeCls}">${_e(d.type)}</span></div>
</div>

<div class="section-title">Détail des travaux / observations</div>
<table>
  <tr><th>Corps d'état</th><th>Désignation</th><th>Localisation</th><th>Remarque</th></tr>
  ${itemRows}
</table>

${d.commentaire ? `<div class="section-title">Commentaire technique</div><p>${_e(d.commentaire).replace(/\n/g,"<br>")}</p>` : ""}

<div class="section-title">Liens photos</div>
<div style="font-size:12px">${photoHtml}</div>

<div class="section-title">Signatures</div>
<div class="sig-area">${sigBoxes}</div>

<div class="footer">Document généré automatiquement par SAVFLOW PRO 3 — ${_e(d.dateStr)}</div>
</body></html>`;
}

/* ============================================================
   DASHBOARD
============================================================ */
function _getDashboard(req) {
  const shR = _sh(S.RECEP);
  if (!shR) return { ok: false, error: "Sheet RECEPTIONS introuvable" };

  const filterProj = _su(req.projectId || "");

  const { headers, rows } = _all(shR);
  const m = _hmap(headers);

  const types   = { SAV:0, LIVRAISON:0, CHANTIER:0 };
  const statuts = {};
  const byCollab = {};
  const byPrest  = {};
  const byCorps  = {};
  let totalWithPhotos = 0;
  let totalRepairMin  = 0; let repairCount = 0;
  let totalTotalMin   = 0; let totalCount  = 0;

  rows.forEach(row => {
    const proj  = _su(row[m["ProjectID"]] || "");
    if (filterProj && filterProj !== "ALL" && proj !== filterProj) return;

    const type  = _su(row[m["Type"]]||"SAV");
    const stat  = _s(row[m["Statut"]]||"");
    const collab = _s(row[m["Collaborateur"]]||row[m["CreatedBy"]]||"-");
    const prest  = _s(row[m["Prestataire"]]||"-");
    const photos = _s(row[m["PhotoLinks"]]||"");
    const createdAt = row[m["CreatedAt"]];
    const startAt   = row[m["StartAt"]];
    const doneAt    = row[m["DoneAt"]];
    const plannedAt = row[m["PlannedAt"]];

    if (types[type] !== undefined) types[type]++;
    else types[type] = (types[type]||0) + 1;
    statuts[stat] = (statuts[stat]||0) + 1;
    if (photos.trim()) totalWithPhotos++;

    // collab
    if (!byCollab[collab]) byCollab[collab] = { total:0, done:0, repMin:0, repCount:0 };
    byCollab[collab].total++;
    if (stat === "Terminée") byCollab[collab].done++;

    const rm = _diffMin(startAt, doneAt);
    if (rm !== null) { byCollab[collab].repMin += rm; byCollab[collab].repCount++; totalRepairMin += rm; repairCount++; }

    const tm = _diffMin(createdAt, doneAt);
    if (tm !== null) { totalTotalMin += tm; totalCount++; }

    // prestataire
    if (prest && prest !== "-") {
      if (!byPrest[prest]) byPrest[prest] = { total:0, done:0, repMin:0, repCount:0, late:0 };
      byPrest[prest].total++;
      if (stat === "Terminée") byPrest[prest].done++;
      if (rm !== null) { byPrest[prest].repMin += rm; byPrest[prest].repCount++; }
      if (doneAt && plannedAt && new Date(doneAt) > new Date(plannedAt)) byPrest[prest].late++;
    }
  });

  // corps d'état from RECEPTION_ITEMS
  const shI = _sh(S.ITEMS);
  if (shI) {
    const { headers: iH, rows: iRows } = _all(shI);
    const iM = _hmap(iH);
    iRows.forEach(r => {
      const corps = _s(r[iM["CorpsEtat"]]||"-");
      if (!byCorps[corps]) byCorps[corps] = 0;
      byCorps[corps]++;
    });
  }

  const total = Object.values(types).reduce((a,b)=>a+b,0);
  const pct   = total ? Math.round(totalWithPhotos/total*100) : 0;
  const avgRepair = repairCount ? Math.round(totalRepairMin/repairCount) : null;
  const avgTotal  = totalCount  ? Math.round(totalTotalMin/totalCount)   : null;

  return {
    ok: true,
    total,
    byType:   types,
    byStatut: statuts,
    photoPct: pct,
    avgRepairMin: avgRepair,
    avgTotalMin:  avgTotal,
    byCollab: Object.keys(byCollab).map(k => ({
      name: k,
      total: byCollab[k].total,
      done:  byCollab[k].done,
      backlog: byCollab[k].total - byCollab[k].done,
      avgRepairMin: byCollab[k].repCount ? Math.round(byCollab[k].repMin/byCollab[k].repCount) : null
    })).sort((a,b) => b.total - a.total),
    byPrest: Object.keys(byPrest).map(k => ({
      name: k,
      total: byPrest[k].total,
      done:  byPrest[k].done,
      avgRepairMin: byPrest[k].repCount ? Math.round(byPrest[k].repMin/byPrest[k].repCount) : null,
      latePct: byPrest[k].total ? Math.round(byPrest[k].late/byPrest[k].total*100) : 0
    })).sort((a,b) => b.total - a.total),
    byCorps: Object.keys(byCorps).map(k => ({
      name: k, count: byCorps[k]
    })).sort((a,b) => b.count - a.count)
  };
}

/* ============================================================
   UTIL
============================================================ */
function _uniq(arr) {
  const seen = {}; return arr.filter(v => { const k = String(v); if(seen[k]) return false; seen[k]=true; return true; });
}
