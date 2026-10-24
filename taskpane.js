/* Onglets verticaux v2 : feuilles Excel en liste verticale, dossiers, étiquettes, actions. */
(function () {
  "use strict";

  const DEMO = /[?&]demo\b/.test(location.search);
  const $ = (id) => document.getElementById(id);
  const treeEl = $("tree"), filterEl = $("filter"), tagBar = $("tags"), showHiddenEl = $("showHidden");
  const countEl = $("count"), addBtn = $("add"), menuEl = $("menu"), modalEl = $("modal"), toastEl = $("toast");

  const COLORS = ["#C0392B", "#E67E22", "#F1C40F", "#27AE60", "#16A085", "#2980B9", "#8E44AD", "#7F8C8D"];
  const ICON = {
    chev: '<svg viewBox="0 0 10 10"><path d="M3.5 2l3 3-3 3"/></svg>',
    folder: '<svg viewBox="0 0 16 16"><path d="M1.5 4.5v8a1 1 0 0 0 1 1h11a1 1 0 0 0 1-1V6a1 1 0 0 0-1-1H8L6.5 3.5h-4a1 1 0 0 0-1 1z"/></svg>',
    check: '<svg viewBox="0 0 12 12"><path d="M2.5 6.5l2.5 2.5 4.5-5"/></svg>',
    dash: '<svg viewBox="0 0 12 12"><path d="M3 6h6"/></svg>'
  };

  let api = null;
  let sheets = [];
  let activeId = null;
  let data = emptyData();
  let sel = new Set();
  let anchor = null;
  let tagFilter = new Set();
  let busy = false;
  let sig = "";
  let flat = [];
  let drag = null;
  let pollTimer = null;
  let saveTimer = null;

  /* ---------- utilitaires ---------- */

  function emptyData() { return { v: 1, folders: [], tags: [], sheets: {} }; }
  function uid(p) { return p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
  function store(k, v) {
    try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; }
  }
  let toastTimer = null;
  function toast(msg) {
    toastEl.textContent = msg;
    toastEl.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toastEl.hidden = true; }, 3200);
  }
  function errMsg(e) {
    const code = e && e.code ? e.code : "";
    if (code === "InvalidArgument" || code === "InvalidOperation") return "Opération refusée par Excel (nom invalide, déjà utilisé, ou classeur protégé).";
    if (code === "AccessDenied") return "Accès refusé (classeur ou structure protégés).";
    return "Erreur : " + (e && e.message ? e.message : e);
  }
  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }
  function plural(n, one, many) { return n + " " + (n > 1 ? many : one); }

  /* ---------- accès Excel ---------- */

  function xrun(fn) { return Excel.run(fn).catch((e) => { toast(errMsg(e)); throw e; }); }

  const realApi = {
    async list() {
      let out;
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets;
        ws.load("items/id,items/name,items/position,items/visibility,items/tabColor");
        const a = ws.getActiveWorksheet();
        a.load("id");
        await ctx.sync();
        out = {
          sheets: ws.items.map((s) => ({ id: s.id, name: s.name, position: s.position, visibility: s.visibility, tabColor: s.tabColor })),
          activeId: a.id
        };
      });
      return out;
    },
    activate(id, unhide) {
      return xrun(async (ctx) => {
        const w = ctx.workbook.worksheets.getItem(id);
        if (unhide) w.visibility = "Visible";
        w.activate();
        await ctx.sync();
      });
    },
    async add(pos) {
      let id;
      await xrun(async (ctx) => {
        const w = ctx.workbook.worksheets.add();
        w.load("id");
        await ctx.sync();
        if (pos != null) w.position = pos;
        w.activate();
        await ctx.sync();
        id = w.id;
      });
      return id;
    },
    rename(id, name) {
      return xrun(async (ctx) => { ctx.workbook.worksheets.getItem(id).name = name; await ctx.sync(); });
    },
    setVisibility(ids, vis, nextId) {
      return xrun(async (ctx) => {
        const ws = ctx.workbook.worksheets;
        if (nextId) ws.getItem(nextId).activate();
        ids.forEach((id) => { ws.getItem(id).visibility = vis; });
        await ctx.sync();
      });
    },
    setColor(ids, c) {
      return xrun(async (ctx) => {
        ids.forEach((id) => { ctx.workbook.worksheets.getItem(id).tabColor = c; });
        await ctx.sync();
      });
    },
    reorder(ids) {
      return xrun(async (ctx) => {
        ids.forEach((id, i) => { ctx.workbook.worksheets.getItem(id).position = i; });
        await ctx.sync();
      });
    },
    async duplicate(id) {
      let nid;
      await xrun(async (ctx) => {
        const src = ctx.workbook.worksheets.getItem(id);
        const c = src.copy("After", src);
        c.load("id");
        c.activate();
        await ctx.sync();
        nid = c.id;
      });
      return nid;
    },
    remove(ids) {
      return xrun(async (ctx) => {
        ids.forEach((id) => ctx.workbook.worksheets.getItem(id).delete());
        await ctx.sync();
      });
    },
    loadData() {
      const v = Office.context.document.settings.get("ov.data");
      if (!v) return null;
      return typeof v === "string" ? JSON.parse(v) : v;
    },
    saveData(d) {
      const s = Office.context.document.settings;
      s.set("ov.data", JSON.stringify(d));
      s.saveAsync(() => {});
    }
  };

  /* ---------- mode démo (hors Excel) ---------- */

  function makeMock() {
    const names = ["Page de garde", "Synthèse", "LOT 01 VRD", "LOT 02 Gros œuvre", "LOT 03 Charpente", "LOT 04 Couverture",
      "LOT 05 Menuiseries ext", "LOT 06 Serrurerie", "LOT 07 Menuiseries int", "LOT 08 Plâtrerie", "LOT 09 Carrelage",
      "LOT 10 Peinture", "Annexes", "Listes"];
    let n = 0;
    const S = names.map((nm, i) => ({ id: "m" + (++n), name: nm, position: i, visibility: i === 13 ? "Hidden" : "Visible", tabColor: i < 2 ? "#2980B9" : "" }));
    let act = "m1";
    const norm = () => { S.sort((a, b) => a.position - b.position).forEach((s, i) => { s.position = i; }); };
    const get = (id) => S.find((s) => s.id === id);
    const ok = () => Promise.resolve();
    return {
      list() { norm(); return Promise.resolve({ sheets: S.map((s) => Object.assign({}, s)), activeId: act }); },
      activate(id, unhide) { if (unhide) get(id).visibility = "Visible"; act = id; return ok(); },
      add(pos) {
        const s = { id: "m" + (++n), name: "Feuil" + n, position: pos == null ? S.length : pos - 0.5, visibility: "Visible", tabColor: "" };
        S.push(s); norm(); act = s.id; return Promise.resolve(s.id);
      },
      rename(id, nm) {
        if (S.some((s) => s.id !== id && s.name.toLowerCase() === nm.toLowerCase())) { toast(errMsg({ code: "InvalidArgument" })); return Promise.reject(new Error("dup")); }
        get(id).name = nm; return ok();
      },
      setVisibility(ids, vis, nextId) { ids.forEach((id) => { get(id).visibility = vis; }); if (nextId) act = nextId; return ok(); },
      setColor(ids, c) { ids.forEach((id) => { get(id).tabColor = c; }); return ok(); },
      reorder(ids) { ids.forEach((id, i) => { get(id).position = i; }); return ok(); },
      duplicate(id) {
        const s = get(id);
        const c = Object.assign({}, s, { id: "m" + (++n), name: s.name + " (2)", position: s.position + 0.5 });
        S.push(c); norm(); act = c.id; return Promise.resolve(c.id);
      },
      remove(ids) {
        ids.forEach((id) => S.splice(S.indexOf(get(id)), 1));
        if (!get(act)) act = S.filter((s) => s.visibility === "Visible")[0].id;
        norm(); return ok();
      },
      loadData() {
        const v = store("ov.demo");
        if (v) return JSON.parse(v);
        const d = emptyData();
        d.folders.push({ id: "f1", name: "Lots techniques", collapsed: false });
        d.tags.push({ id: "t1", name: "À viser", color: "#E67E22" }, { id: "t2", name: "Validé", color: "#27AE60" });
        for (let i = 3; i <= 12; i++) d.sheets["m" + i] = { n: names[i - 1], f: "f1", t: i === 7 || i === 8 ? ["t1"] : i < 6 ? ["t2"] : [] };
        return d;
      },
      saveData(d) { store("ov.demo", JSON.stringify(d)); }
    };
  }

  /* ---------- données (dossiers, étiquettes) ---------- */

  function byId(id) { return sheets.find((s) => s.id === id); }
  function meta(id) {
    if (!data.sheets[id]) data.sheets[id] = { n: (byId(id) || {}).name || "", f: null, t: [] };
    return data.sheets[id];
  }
  function folderOf(id) { return (data.sheets[id] && data.sheets[id].f) || null; }
  function tagsOf(id) { return (data.sheets[id] && data.sheets[id].t) || []; }
  function folderById(fid) { return data.folders.find((f) => f.id === fid); }
  function tagById(tid) { return data.tags.find((t) => t.id === tid); }
  function sheetsIn(fid, list) { return list.filter((s) => folderOf(s.id) === fid); }
  function selOrdered() { return sheets.filter((s) => sel.has(s.id)).map((s) => s.id); }

  function save() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => { try { api.saveData(data); } catch (e) { toast("Impossible d'enregistrer l'organisation."); } }, 250);
  }

  function reconcile() {
    const live = new Set(sheets.map((s) => s.id));
    let ch = false;
    for (const s of sheets) {
      if (!data.sheets[s.id]) {
        const k = Object.keys(data.sheets).find((k2) => !live.has(k2) && data.sheets[k2].n === s.name);
        if (!k) continue;
        data.sheets[s.id] = data.sheets[k];
        delete data.sheets[k];
        ch = true;
      }
      if (data.sheets[s.id].n !== s.name) { data.sheets[s.id].n = s.name; ch = true; }
    }
    for (const k of Object.keys(data.sheets)) if (!live.has(k)) { delete data.sheets[k]; ch = true; }
    const fids = new Set(data.folders.map((f) => f.id));
    const tids = new Set(data.tags.map((t) => t.id));
    for (const k in data.sheets) {
      const e = data.sheets[k];
      if (e.f && !fids.has(e.f)) { e.f = null; ch = true; }
      const t = (e.t || []).filter((x) => tids.has(x));
      if (t.length !== (e.t || []).length) { e.t = t; ch = true; }
    }
    for (const id of Array.from(sel)) if (!live.has(id)) sel.delete(id);
    if (ch) save();
  }

  /* ---------- lecture du classeur ---------- */

  async function refresh(force) {
    if (busy && !force) return;
    let r;
    try { r = await api.list(); } catch (e) { return; }
    sheets = r.sheets.sort((a, b) => a.position - b.position);
    activeId = r.activeId;
    reconcile();
    const s = JSON.stringify([sheets, activeId]);
    if (s !== sig || force) {
      sig = s;
      if (!busy || force) render();
    }
  }

  /* ---------- affichage ---------- */

  function render() {
    renderTags();
    const q = filterEl.value.trim().toLowerCase();
    const filtering = !!q || tagFilter.size > 0;
    const list = sheets.filter((s) =>
      (showHiddenEl.checked || s.visibility === "Visible") &&
      (!q || s.name.toLowerCase().includes(q)) &&
      (!tagFilter.size || tagsOf(s.id).some((t) => tagFilter.has(t))));

    treeEl.textContent = "";
    flat = [];
    const frag = document.createDocumentFragment();
    for (const f of data.folders) {
      const items = sheetsIn(f.id, list);
      if (filtering && !items.length) continue;
      const open = filtering || !f.collapsed;
      frag.appendChild(folderRow(f, items.length, open));
      if (open) for (const s of items) { frag.appendChild(sheetRow(s, true)); flat.push(s.id); }
    }
    const rest = sheetsIn(null, list);
    if (data.folders.length && (rest.length || !filtering)) frag.appendChild(sectionRow());
    for (const s of rest) { frag.appendChild(sheetRow(s, false)); flat.push(s.id); }
    if (!flat.length && !frag.childNodes.length) frag.appendChild(el("div", "empty", filtering ? "Aucune feuille ne correspond." : "Aucune feuille."));
    treeEl.appendChild(frag);
    updateCount();
    const act = treeEl.querySelector(".row.active");
    if (act) act.scrollIntoView({ block: "nearest" });
  }

  function sheetRow(s, inFolder) {
    const r = el("div", "row sheet");
    r.dataset.id = s.id;
    r.draggable = true;
    r.title = s.name;
    if (inFolder) r.style.setProperty("--indent", "15px");
    if (s.id === activeId) r.classList.add("active");
    if (sel.has(s.id)) r.classList.add("sel");
    if (s.visibility !== "Visible") r.classList.add("hidden-sheet");
    const c = el("span", "tabc");
    if (s.tabColor) c.style.background = s.tabColor;
    r.appendChild(c);
    r.appendChild(el("span", "name", s.name));
    const tg = tagsOf(s.id).map(tagById).filter(Boolean);
    if (tg.length) {
      const d = el("span", "tdots");
      tg.forEach((t) => { const i = el("i"); i.style.background = t.color; i.title = t.name; d.appendChild(i); });
      r.appendChild(d);
    }
    if (s.visibility === "Hidden") r.appendChild(el("span", "badge", "masquée"));
    if (s.visibility === "VeryHidden") r.appendChild(el("span", "badge", "très masquée"));
    return r;
  }

  function folderRow(f, count, open) {
    const r = el("div", "row folder" + (open ? " open" : ""));
    r.dataset.fid = f.id;
    r.draggable = true;
    const ch = el("span", "chev"); ch.innerHTML = ICON.chev; r.appendChild(ch);
    const fi = el("span", "ficon"); fi.innerHTML = ICON.folder; r.appendChild(fi);
    r.appendChild(el("span", "name", f.name));
    r.appendChild(el("span", "cnt", String(count)));
    return r;
  }

  function sectionRow() {
    const r = el("div", "row section", "Sans dossier");
    r.dataset.section = "1";
    return r;
  }

  function renderTags() {
    tagBar.textContent = "";
    if (!data.tags.length) { tagBar.hidden = true; return; }
    tagBar.hidden = false;
    for (const t of data.tags) {
      const c = el("button", "chip" + (tagFilter.has(t.id) ? " on" : ""));
      c.dataset.tid = t.id;
      c.title = "Filtrer sur cette étiquette (clic droit pour la modifier)";
      const d = el("span", "dot"); d.style.background = t.color; c.appendChild(d);
      c.appendChild(el("span", null, t.name));
      tagBar.appendChild(c);
    }
    if (tagFilter.size) {
      const c = el("button", "chip clear", "Tout afficher");
      c.dataset.clear = "1";
      tagBar.appendChild(c);
    }
  }

  function markSel() {
    treeEl.querySelectorAll(".row.sheet").forEach((r) => r.classList.toggle("sel", sel.has(r.dataset.id)));
    updateCount();
  }

  function updateCount() {
    if (sel.size > 1) { countEl.textContent = plural(sel.size, "sélectionnée", "sélectionnées"); return; }
    const nh = sheets.filter((s) => s.visibility !== "Visible").length;
    countEl.textContent = plural(sheets.length, "feuille", "feuilles") + (nh ? " · " + plural(nh, "masquée", "masquées") : "");
  }

  /* ---------- actions sur les feuilles ---------- */

  async function activate(id) {
    const s = byId(id);
    if (!s) return;
    if (s.visibility === "VeryHidden") { toast("Feuille très masquée : affichez-la d'abord (clic droit)."); return; }
    activeId = id;
    render();
    await api.activate(id, s.visibility === "Hidden").catch(() => {});
    refresh(true);
  }

  async function newSheet(fid) {
    let pos = null;
    if (fid) {
      const inF = sheetsIn(fid, sheets);
      if (inF.length) pos = inF[inF.length - 1].position + 1;
    }
    let id;
    try { id = await api.add(pos); } catch (e) { return; }
    if (fid) { const f = folderById(fid); if (f) f.collapsed = false; meta(id).f = fid; save(); }
    filterEl.value = "";
    sel = new Set([id]); anchor = id;
    await refresh(true);
    if (fid) { meta(id).f = fid; save(); render(); }
    startRename(id);
  }

  async function renameSheet(id, name) {
    name = name.trim();
    const s = byId(id);
    if (!s || !name || name === s.name) return;
    if (/[\\\/\?\*\[\]:]/.test(name)) { toast("Caractères interdits : \\ / ? * [ ] :"); return; }
    if (name.length > 31) { toast("31 caractères maximum."); return; }
    if (/^'|'$/.test(name)) { toast("Le nom ne peut pas commencer ou finir par une apostrophe."); return; }
    await api.rename(id, name).catch(() => {});
  }

  async function duplicate(id) {
    let nid;
    try { nid = await api.duplicate(id); } catch (e) { return; }
    const e = data.sheets[id];
    if (e && nid) { data.sheets[nid] = { n: "", f: e.f, t: (e.t || []).slice() }; save(); }
    sel = new Set([nid]); anchor = nid;
    refresh(true);
  }

  async function toggleHidden(ids) {
    const targets = ids.map(byId).filter(Boolean);
    const vis = targets.filter((s) => s.visibility === "Visible");
    if (vis.length) {
      const remaining = sheets.filter((s) => s.visibility === "Visible" && !ids.includes(s.id));
      if (!remaining.length) { toast("Le classeur doit garder au moins une feuille visible."); return; }
      let next = null;
      if (ids.includes(activeId)) {
        const a = byId(activeId);
        next = (remaining.find((s) => s.position > a.position) || remaining[remaining.length - 1]).id;
      }
      await api.setVisibility(vis.map((s) => s.id), "Hidden", next).catch(() => {});
      if (!showHiddenEl.checked) toast(plural(vis.length, "feuille masquée", "feuilles masquées") + ". Cochez « Feuilles masquées » pour les revoir.");
    } else {
      await api.setVisibility(targets.map((s) => s.id), "Visible").catch(() => {});
    }
    refresh(true);
  }

  async function setColor(ids, c) {
    await api.setColor(ids, c).catch(() => {});
    refresh(true);
  }

  async function deleteSheets(ids) {
    ids = ids.filter((id) => byId(id));
    if (!ids.length) return;
    const remaining = sheets.filter((s) => s.visibility === "Visible" && !ids.includes(s.id));
    if (!remaining.length) { toast("Le classeur doit garder au moins une feuille visible."); return; }
    const names = ids.map((id) => byId(id).name);
    const many = ids.length > 1;
    const ok = await ask({
      title: many ? "Supprimer " + ids.length + " feuilles ?" : "Supprimer la feuille ?",
      text: (many ? names.slice(0, 6).join(", ") + (names.length > 6 ? "…" : "") : "« " + names[0] + " »") +
        (many ? " seront supprimées avec leur contenu" : " sera supprimée avec son contenu") + ". Cette action ne peut pas être annulée.",
      ok: "Supprimer",
      danger: true
    });
    if (!ok) return;
    try { await api.remove(ids); } catch (e) { refresh(true); return; }
    ids.forEach((id) => { delete data.sheets[id]; sel.delete(id); });
    save();
    refresh(true);
  }

  async function reorderSheets(ids, targetId, after) {
    const cur = sheets.map((s) => s.id);
    const order = cur.filter((id) => !ids.includes(id));
    let i = order.indexOf(targetId);
    if (i < 0) return;
    if (after) i++;
    order.splice.apply(order, [i, 0].concat(ids));
    if (order.join("|") === cur.join("|")) return;
    await api.reorder(order).catch(() => {});
    refresh(true);
  }

  async function arrange() {
    if (!data.folders.length) return;
    const order = [];
    data.folders.forEach((f) => sheetsIn(f.id, sheets).forEach((s) => order.push(s.id)));
    sheetsIn(null, sheets).forEach((s) => order.push(s.id));
    await api.reorder(order).catch(() => {});
    await refresh(true);
    toast("Ordre des feuilles du classeur aligné sur les dossiers.");
  }

  /* ---------- dossiers ---------- */

  function moveToFolder(ids, fid) {
    ids.forEach((id) => { meta(id).f = fid; });
    if (fid) { const f = folderById(fid); if (f) f.collapsed = false; }
    save(); render();
  }

  async function newFolder(ids) {
    const name = await ask({ title: "Nouveau dossier", input: true, value: "", placeholder: "Nom du dossier", ok: "Créer" });
    if (!name) return;
    const f = { id: uid("f"), name: name, collapsed: false };
    data.folders.push(f);
    if (ids) ids.forEach((id) => { meta(id).f = f.id; });
    save(); render();
  }

  async function renameFolder(fid) {
    const f = folderById(fid);
    if (!f) return;
    const name = await ask({ title: "Renommer le dossier", input: true, value: f.name, ok: "Renommer" });
    if (!name) return;
    f.name = name; save(); render();
  }

  async function deleteFolder(fid) {
    const f = folderById(fid);
    if (!f) return;
    const ok = await ask({ title: "Supprimer le dossier ?", text: "Le dossier « " + f.name + " » sera retiré. Ses feuilles sont conservées et repassent « Sans dossier ».", ok: "Supprimer", danger: true });
    if (!ok) return;
    data.folders = data.folders.filter((x) => x.id !== fid);
    Object.keys(data.sheets).forEach((k) => { if (data.sheets[k].f === fid) data.sheets[k].f = null; });
    save(); render();
  }

  function toggleFolder(fid) {
    const f = folderById(fid);
    if (!f) return;
    f.collapsed = !f.collapsed; save(); render();
  }

  function setAllCollapsed(v) { data.folders.forEach((f) => { f.collapsed = v; }); save(); render(); }

  function selectFolder(fid) {
    const f = folderById(fid);
    if (f) f.collapsed = false;
    sel = new Set(sheetsIn(fid, sheets).filter((s) => showHiddenEl.checked || s.visibility === "Visible").map((s) => s.id));
    save(); render();
  }

  /* ---------- étiquettes ---------- */

  function toggleTag(ids, tid) {
    const all = ids.every((id) => tagsOf(id).includes(tid));
    ids.forEach((id) => {
      const e = meta(id);
      e.t = (e.t || []).filter((x) => x !== tid);
      if (!all) e.t.push(tid);
    });
    save(); render();
  }

  async function newTag(ids) {
    const name = await ask({ title: "Nouvelle étiquette", input: true, value: "", placeholder: "Ex. À viser, Validé, En attente", ok: "Créer" });
    if (!name) return;
    const t = { id: uid("t"), name: name, color: COLORS[data.tags.length % COLORS.length] };
    data.tags.push(t);
    if (ids) ids.forEach((id) => { const e = meta(id); e.t = (e.t || []).concat(t.id); });
    save(); render();
  }

  async function renameTag(tid) {
    const t = tagById(tid);
    if (!t) return;
    const name = await ask({ title: "Renommer l'étiquette", input: true, value: t.name, ok: "Renommer" });
    if (!name) return;
    t.name = name; save(); render();
  }

  async function deleteTag(tid) {
    const t = tagById(tid);
    if (!t) return;
    const ok = await ask({ title: "Supprimer l'étiquette ?", text: "L'étiquette « " + t.name + " » sera retirée de toutes les feuilles.", ok: "Supprimer", danger: true });
    if (!ok) return;
    data.tags = data.tags.filter((x) => x.id !== tid);
    Object.keys(data.sheets).forEach((k) => { data.sheets[k].t = (data.sheets[k].t || []).filter((x) => x !== tid); });
    tagFilter.delete(tid);
    save(); render();
  }

  /* ---------- renommage en place ---------- */

  function startRename(id) {
    const row = treeEl.querySelector('.row.sheet[data-id="' + CSS.escape(id) + '"]');
    const s = byId(id);
    if (!row || !s) return;
    busy = true;
    row.draggable = false;
    const nameEl = row.querySelector(".name");
    const input = el("input", "rename");
    input.value = s.name;
    input.maxLength = 31;
    nameEl.replaceWith(input);
    input.focus();
    input.select();
    let done = false;
    const finish = async (commit) => {
      if (done) return;
      done = true;
      if (commit) await renameSheet(id, input.value);
      busy = false;
      await refresh(true);
      treeEl.focus({ preventScroll: true });
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); finish(true); }
      else if (e.key === "Escape") { e.preventDefault(); finish(false); }
      e.stopPropagation();
    });
    input.addEventListener("blur", () => finish(true));
    ["click", "dblclick", "mousedown"].forEach((t) => input.addEventListener(t, (e) => e.stopPropagation()));
  }

  /* ---------- menus ---------- */

  const menu = { stack: [], x: 0, y: 0 };

  function openMenu(x, y, builder) {
    menu.stack = [builder]; menu.x = x; menu.y = y;
    menuEl.hidden = false;
    drawMenu();
  }
  function closeMenu() { menuEl.hidden = true; menu.stack = []; menuEl.textContent = ""; }

  function drawMenu() {
    const b = menu.stack[menu.stack.length - 1];
    menuEl.textContent = "";
    const m = {
      item(label, fn, o) {
        o = o || {};
        const bt = el("button", "mi" + (o.danger ? " danger" : "") + (o.disabled ? " disabled" : "") + (o.cls ? " " + o.cls : ""));
        const ic = el("span", "mic");
        if (o.icon) ic.innerHTML = o.icon;
        bt.appendChild(ic);
        if (o.dot) { const d = el("span", "mdot"); d.style.background = o.dot; bt.appendChild(d); }
        bt.appendChild(el("span", "ml", label));
        if (o.hint) bt.appendChild(el("span", "mk", o.hint));
        if (o.sub) bt.appendChild(el("span", "mk", "›"));
        bt.addEventListener("click", (e) => {
          e.stopPropagation();
          if (o.disabled) return;
          if (o.sub) { menu.stack.push(o.sub); drawMenu(); return; }
          if (o.keep) { if (fn) fn(); drawMenu(); return; }
          closeMenu();
          if (fn) fn();
        });
        menuEl.appendChild(bt);
      },
      sep() { menuEl.appendChild(el("div", "msep")); },
      label(t) { menuEl.appendChild(el("div", "mlabel", t)); },
      swatches(cur, fn, allowNone) {
        const w = el("div", "swatches");
        const opts = (allowNone === false ? [] : [""]).concat(COLORS);
        opts.forEach((c) => {
          const s = el("button", c ? "" : "none");
          if (c) s.style.background = c;
          s.title = c || "Aucune";
          if ((cur || "").toUpperCase() === c.toUpperCase()) s.classList.add("cur");
          s.addEventListener("click", (e) => { e.stopPropagation(); closeMenu(); fn(c); });
          w.appendChild(s);
        });
        menuEl.appendChild(w);
      }
    };
    if (menu.stack.length > 1) {
      m.item("Retour", () => { menu.stack.pop(); }, { keep: true, icon: '<svg viewBox="0 0 12 12"><path d="M7.5 2.5L4 6l3.5 3.5"/></svg>', cls: "back" });
      m.sep();
    }
    b(m);
    const r = menuEl.getBoundingClientRect();
    menuEl.style.left = Math.max(4, Math.min(menu.x, window.innerWidth - r.width - 4)) + "px";
    menuEl.style.top = Math.max(4, Math.min(menu.y, window.innerHeight - r.height - 4)) + "px";
  }

  function sheetMenu(ids) {
    const one = ids.length === 1;
    return (m) => {
      if (!one) m.label(plural(ids.length, "feuille sélectionnée", "feuilles sélectionnées"));
      if (one) m.item("Renommer", () => startRename(ids[0]), { hint: "F2" });
      if (one) m.item("Dupliquer", () => duplicate(ids[0]));
      m.item("Déplacer vers", null, { sub: folderPick(ids) });
      m.item("Étiquettes", null, { sub: tagPick(ids) });
      m.item("Couleur d'onglet", null, { sub: (mm) => mm.swatches(one ? byId(ids[0]).tabColor : "", (c) => setColor(ids, c)) });
      const anyVis = ids.some((id) => byId(id) && byId(id).visibility === "Visible");
      m.item(anyVis ? "Masquer" : "Afficher", () => toggleHidden(ids));
      m.sep();
      m.item(one ? "Supprimer la feuille" : "Supprimer les feuilles", () => deleteSheets(ids), { danger: true, hint: "⌫" });
    };
  }

  function folderPick(ids) {
    return (m) => {
      const fs = new Set(ids.map(folderOf));
      const cur = fs.size === 1 ? Array.from(fs)[0] : undefined;
      if (!data.folders.length) m.label("Aucun dossier pour l'instant");
      data.folders.forEach((f) => m.item(f.name, () => moveToFolder(ids, f.id), { icon: cur === f.id ? ICON.check : null }));
      if (data.folders.length) m.item("Sans dossier", () => moveToFolder(ids, null), { icon: cur === null ? ICON.check : null });
      m.sep();
      m.item("Nouveau dossier…", () => newFolder(ids));
    };
  }

  function tagPick(ids) {
    return (m) => {
      if (!data.tags.length) m.label("Aucune étiquette pour l'instant");
      data.tags.forEach((t) => {
        const n = ids.filter((id) => tagsOf(id).includes(t.id)).length;
        m.item(t.name, () => toggleTag(ids, t.id), { keep: true, dot: t.color, icon: n === ids.length ? ICON.check : n ? ICON.dash : null });
      });
      m.sep();
      m.item("Nouvelle étiquette…", () => newTag(ids));
    };
  }

  function folderMenu(fid) {
    const f = folderById(fid);
    return (m) => {
      m.item("Renommer", () => renameFolder(fid));
      m.item("Nouvelle feuille ici", () => newSheet(fid));
      m.item("Sélectionner ses feuilles", () => selectFolder(fid));
      m.item(f.collapsed ? "Déplier" : "Replier", () => toggleFolder(fid));
      m.sep();
      m.item("Ranger le classeur selon les dossiers", arrange);
      m.sep();
      m.item("Supprimer le dossier", () => deleteFolder(fid), { danger: true });
    };
  }

  function addMenu() {
    const nf = !data.folders.length;
    return (m) => {
      m.item("Nouvelle feuille", () => newSheet(null));
      m.item("Nouveau dossier…", () => newFolder(sel.size > 1 ? selOrdered() : null));
      m.item("Nouvelle étiquette…", () => newTag());
      m.sep();
      m.item("Tout replier", () => setAllCollapsed(true), { disabled: nf });
      m.item("Tout déplier", () => setAllCollapsed(false), { disabled: nf });
      m.item("Ranger le classeur selon les dossiers", arrange, { disabled: nf });
    };
  }

  function tagMenu(tid) {
    const t = tagById(tid);
    return (m) => {
      m.item("Renommer…", () => renameTag(tid));
      m.item("Couleur", null, { sub: (mm) => mm.swatches(t.color, (c) => { t.color = c; save(); render(); }, false) });
      m.item("Sélectionner ses feuilles", () => {
        sel = new Set(sheets.filter((s) => tagsOf(s.id).includes(tid)).map((s) => s.id));
        render();
      });
      m.sep();
      m.item("Supprimer l'étiquette", () => deleteTag(tid), { danger: true });
    };
  }

  /* ---------- boîte de dialogue ---------- */

  function ask(o) {
    return new Promise((resolve) => {
      closeMenu();
      modalEl.textContent = "";
      const d = el("div", "dlg");
      d.appendChild(el("h3", null, o.title));
      if (o.text) d.appendChild(el("p", null, o.text));
      let inp = null;
      if (o.input) {
        inp = el("input");
        inp.value = o.value || "";
        if (o.placeholder) inp.placeholder = o.placeholder;
        d.appendChild(inp);
      }
      const btns = el("div", "btns");
      const cancel = el("button", "btn", "Annuler");
      const okb = el("button", "btn " + (o.danger ? "danger" : "primary"), o.ok || "OK");
      btns.appendChild(cancel); btns.appendChild(okb);
      d.appendChild(btns);
      modalEl.appendChild(d);
      modalEl.hidden = false;
      const done = (v) => { modalEl.hidden = true; modalEl.textContent = ""; resolve(v); treeEl.focus({ preventScroll: true }); };
      const accept = () => {
        if (inp) { const v = inp.value.trim(); if (!v) { inp.focus(); return; } done(v); } else done(true);
      };
      cancel.addEventListener("click", () => done(null));
      okb.addEventListener("click", accept);
      modalEl.onmousedown = (e) => { if (e.target === modalEl) done(null); };
      d.addEventListener("keydown", (e) => {
        if (e.key === "Enter") { e.preventDefault(); accept(); }
        else if (e.key === "Escape") { e.preventDefault(); done(null); }
        e.stopPropagation();
      });
      setTimeout(() => { if (inp) { inp.focus(); inp.select(); } else okb.focus(); }, 0);
    });
  }

  /* ---------- interactions ---------- */

  function bindUi() {
    treeEl.addEventListener("click", (e) => {
      const row = e.target.closest(".row");
      if (!row || busy) return;
      if (row.classList.contains("folder")) { toggleFolder(row.dataset.fid); return; }
      if (!row.classList.contains("sheet")) return;
      const id = row.dataset.id;
      treeEl.focus({ preventScroll: true });
      if (e.metaKey || e.ctrlKey) {
        if (sel.has(id)) sel.delete(id); else sel.add(id);
        anchor = id; markSel(); return;
      }
      if (e.shiftKey && anchor && flat.includes(anchor)) {
        const a = flat.indexOf(anchor), b = flat.indexOf(id);
        sel = new Set(flat.slice(Math.min(a, b), Math.max(a, b) + 1));
        markSel(); return;
      }
      sel = new Set([id]); anchor = id;
      activate(id);
    });

    treeEl.addEventListener("dblclick", (e) => {
      const row = e.target.closest(".row");
      if (!row) return;
      if (row.classList.contains("sheet")) startRename(row.dataset.id);
      else if (row.classList.contains("folder")) { toggleFolder(row.dataset.fid); renameFolder(row.dataset.fid); }
    });

    treeEl.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      if (busy) return;
      const row = e.target.closest(".row");
      if (row && row.classList.contains("folder")) return openMenu(e.clientX, e.clientY, folderMenu(row.dataset.fid));
      if (row && row.classList.contains("sheet")) {
        const id = row.dataset.id;
        if (!sel.has(id)) { sel = new Set([id]); anchor = id; markSel(); }
        return openMenu(e.clientX, e.clientY, sheetMenu(selOrdered()));
      }
      openMenu(e.clientX, e.clientY, addMenu());
    });

    treeEl.addEventListener("keydown", (e) => {
      if (busy || !modalEl.hidden) return;
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        if (!flat.length) return;
        let i = flat.indexOf(activeId);
        i = i < 0 ? 0 : Math.max(0, Math.min(flat.length - 1, i + (e.key === "ArrowDown" ? 1 : -1)));
        const id = flat[i];
        sel = new Set([id]); anchor = id;
        activate(id);
      } else if (e.key === "F2") {
        e.preventDefault();
        startRename(sel.size === 1 ? Array.from(sel)[0] : activeId);
      } else if (e.key === "Delete" || e.key === "Backspace") {
        e.preventDefault();
        deleteSheets(sel.size ? selOrdered() : [activeId]);
      } else if (e.key === "Escape") {
        sel.clear(); markSel();
      } else if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "f") {
        e.preventDefault(); filterEl.focus(); filterEl.select();
      }
    });

    tagBar.addEventListener("click", (e) => {
      const c = e.target.closest(".chip");
      if (!c) return;
      if (c.dataset.clear) tagFilter.clear();
      else if (tagFilter.has(c.dataset.tid)) tagFilter.delete(c.dataset.tid);
      else tagFilter.add(c.dataset.tid);
      render();
    });
    tagBar.addEventListener("contextmenu", (e) => {
      const c = e.target.closest(".chip");
      e.preventDefault();
      if (c && c.dataset.tid) openMenu(e.clientX, e.clientY, tagMenu(c.dataset.tid));
    });

    addBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      if (!menuEl.hidden) { closeMenu(); return; }
      const r = addBtn.getBoundingClientRect();
      openMenu(r.right - 190, r.bottom + 4, addMenu());
    });

    document.addEventListener("mousedown", (e) => { if (!menuEl.hidden && !menuEl.contains(e.target) && e.target !== addBtn && !addBtn.contains(e.target)) closeMenu(); });
    document.addEventListener("keydown", (e) => { if (e.key === "Escape" && !menuEl.hidden) closeMenu(); });
    window.addEventListener("blur", closeMenu);

    filterEl.addEventListener("input", render);
    filterEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { if (flat[0]) { sel = new Set([flat[0]]); activate(flat[0]); } }
      else if (e.key === "Escape") { filterEl.value = ""; render(); treeEl.focus(); }
      else if (e.key === "ArrowDown") { e.preventDefault(); treeEl.focus(); }
    });

    showHiddenEl.checked = store("ov.showHidden") === "1";
    showHiddenEl.addEventListener("change", () => { store("ov.showHidden", showHiddenEl.checked ? "1" : "0"); render(); });

    /* glisser-déposer */
    treeEl.addEventListener("dragstart", (e) => {
      const row = e.target.closest(".row");
      if (!row || row.classList.contains("section")) { e.preventDefault(); return; }
      if (row.classList.contains("sheet")) {
        const id = row.dataset.id;
        if (!sel.has(id)) { sel = new Set([id]); anchor = id; markSel(); }
        drag = { kind: "sheet", ids: selOrdered() };
        treeEl.querySelectorAll(".row.sheet.sel").forEach((r) => r.classList.add("dragging"));
      } else {
        drag = { kind: "folder", fid: row.dataset.fid };
        row.classList.add("dragging");
      }
      busy = true;
      e.dataTransfer.effectAllowed = "move";
      try { e.dataTransfer.setData("text/plain", "ov"); } catch (_) {}
    });
    treeEl.addEventListener("dragover", (e) => {
      if (!drag) return;
      const row = e.target.closest(".row");
      clearMarks();
      if (!row) return;
      const r = row.getBoundingClientRect();
      const after = e.clientY > r.top + r.height / 2;
      let mark = null;
      if (drag.kind === "sheet") {
        if (row.classList.contains("folder") || row.classList.contains("section")) mark = "drop-into";
        else if (row.classList.contains("sheet") && !drag.ids.includes(row.dataset.id)) mark = after ? "drop-after" : "drop-before";
      } else if (row.classList.contains("folder") && row.dataset.fid !== drag.fid) {
        mark = after ? "drop-after" : "drop-before";
      }
      if (mark) { e.preventDefault(); e.dataTransfer.dropEffect = "move"; row.classList.add(mark); }
    });
    treeEl.addEventListener("drop", (e) => {
      if (!drag) return;
      e.preventDefault();
      const row = treeEl.querySelector(".drop-into,.drop-before,.drop-after");
      const d = drag;
      const how = row ? (row.classList.contains("drop-into") ? "into" : row.classList.contains("drop-after") ? "after" : "before") : null;
      endDrag();
      if (!row) { render(); return; }
      if (d.kind === "sheet") {
        if (how === "into") { moveToFolder(d.ids, row.dataset.section ? null : row.dataset.fid); return; }
        const tid = row.dataset.id;
        const tf = folderOf(tid);
        d.ids.forEach((id) => { meta(id).f = tf; });
        save(); render();
        reorderSheets(d.ids, tid, how === "after");
      } else {
        const fi = data.folders.findIndex((f) => f.id === d.fid);
        const moved = data.folders.splice(fi, 1)[0];
        let ti = data.folders.findIndex((f) => f.id === row.dataset.fid);
        if (how === "after") ti++;
        data.folders.splice(ti, 0, moved);
        save(); render();
      }
    });
    treeEl.addEventListener("dragend", () => { if (drag) { endDrag(); render(); } });
  }

  function clearMarks() {
    treeEl.querySelectorAll(".drop-before,.drop-after,.drop-into").forEach((r) => r.classList.remove("drop-before", "drop-after", "drop-into"));
  }
  function endDrag() {
    clearMarks();
    treeEl.querySelectorAll(".dragging").forEach((r) => r.classList.remove("dragging"));
    drag = null;
    busy = false;
  }

  /* ---------- événements Excel ---------- */

  async function bindExcelEvents() {
    const handler = () => { refresh(); };
    try {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets;
        ["onActivated", "onAdded", "onDeleted", "onNameChanged", "onVisibilityChanged", "onMoved"].forEach((n) => {
          try { if (ws[n]) ws[n].add(handler); } catch (_) {}
        });
        await ctx.sync();
      });
    } catch (_) {}
  }

  function startPolling() {
    clearInterval(pollTimer);
    pollTimer = setInterval(() => { if (!document.hidden) refresh(); }, 1500);
  }

  /* ---------- démarrage ---------- */

  function init() {
    try {
      const d = api.loadData();
      if (d && Array.isArray(d.folders)) data = Object.assign(emptyData(), d);
    } catch (e) {}
    bindUi();
    refresh(true);
    if (!DEMO) bindExcelEvents();
    startPolling();
  }

  Office.onReady((info) => {
    if (DEMO) { api = makeMock(); init(); return; }
    if (info.host !== Office.HostType.Excel) {
      treeEl.appendChild(el("div", "empty", "Ce complément fonctionne dans Excel."));
      return;
    }
    api = realApi;
    init();
  });
})();
