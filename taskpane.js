/* Onglets verticaux : panneau latéral listant les feuilles Excel verticalement. */
(function () {
  "use strict";

  const listEl = document.getElementById("list");
  const filterEl = document.getElementById("filter");
  const showHiddenEl = document.getElementById("showHidden");
  const countEl = document.getElementById("count");
  const addBtn = document.getElementById("add");
  const menuEl = document.getElementById("menu");
  const toastEl = document.getElementById("toast");

  let sheets = [];
  let activeId = null;
  let signature = "";
  let busy = false;
  let menuSheetId = null;
  let dragId = null;
  let pollTimer = null;

  const PREF_HIDDEN = "ov.showHidden";

  function store(key, val) {
    try { if (val === undefined) return localStorage.getItem(key); localStorage.setItem(key, val); } catch (e) { return null; }
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

  function run(fn) {
    return Excel.run(fn).catch((e) => { toast(errMsg(e)); throw e; });
  }

  async function refresh(force) {
    if (busy && !force) return;
    try {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets;
        ws.load("items/id,items/name,items/position,items/visibility,items/tabColor");
        const act = ws.getActiveWorksheet();
        act.load("id");
        await ctx.sync();
        sheets = ws.items
          .map((s) => ({ id: s.id, name: s.name, position: s.position, visibility: s.visibility, tabColor: s.tabColor }))
          .sort((a, b) => a.position - b.position);
        activeId = act.id;
      });
    } catch (e) {
      return;
    }
    const sig = JSON.stringify([sheets, activeId]);
    if (sig !== signature || force) {
      signature = sig;
      if (!busy || force) render();
    }
  }

  function render() {
    const q = filterEl.value.trim().toLowerCase();
    const showHidden = showHiddenEl.checked;
    const visible = sheets.filter((s) => (showHidden || s.visibility === "Visible") && (!q || s.name.toLowerCase().includes(q)));

    listEl.textContent = "";
    for (const s of visible) {
      const li = document.createElement("li");
      li.className = "item";
      li.dataset.id = s.id;
      li.setAttribute("role", "option");
      li.draggable = !q;
      if (s.id === activeId) { li.classList.add("active"); li.setAttribute("aria-selected", "true"); }
      if (s.visibility !== "Visible") li.classList.add("hidden-sheet");

      if (s.tabColor) {
        const c = document.createElement("span");
        c.className = "tab-color";
        c.style.background = s.tabColor;
        li.appendChild(c);
      }
      const n = document.createElement("span");
      n.className = "name";
      n.textContent = s.name;
      n.title = s.name;
      li.appendChild(n);

      if (s.visibility === "Hidden") li.appendChild(badge("masquée"));
      if (s.visibility === "VeryHidden") li.appendChild(badge("très masquée"));

      listEl.appendChild(li);
    }

    if (!visible.length) {
      const e = document.createElement("li");
      e.className = "empty";
      e.textContent = q ? "Aucune feuille ne correspond." : "Aucune feuille.";
      listEl.appendChild(e);
    }

    const nHidden = sheets.filter((s) => s.visibility !== "Visible").length;
    countEl.textContent = sheets.length + (sheets.length > 1 ? " feuilles" : " feuille") + (nHidden ? " · " + nHidden + " masquée" + (nHidden > 1 ? "s" : "") : "");

    const act = listEl.querySelector(".item.active");
    if (act) act.scrollIntoView({ block: "nearest" });
  }

  function badge(txt) {
    const b = document.createElement("span");
    b.className = "badge";
    b.textContent = txt;
    return b;
  }

  function byId(id) { return sheets.find((s) => s.id === id); }

  async function activate(id) {
    const s = byId(id);
    if (!s) return;
    if (s.visibility === "VeryHidden") { toast("Feuille très masquée : affichez-la d'abord avec le clic droit."); return; }
    activeId = id; render();
    await run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(id);
      if (s.visibility === "Hidden") ws.visibility = "Visible";
      ws.activate();
      await ctx.sync();
    }).catch(() => {});
    refresh(true);
  }

  async function addSheet() {
    await run(async (ctx) => {
      const ws = ctx.workbook.worksheets.add();
      ws.activate();
      await ctx.sync();
    }).catch(() => {});
    filterEl.value = "";
    refresh(true);
  }

  async function renameSheet(id, newName) {
    newName = newName.trim();
    const s = byId(id);
    if (!s || !newName || newName === s.name) return;
    if (/[\\\/\?\*\[\]:]/.test(newName)) { toast("Caractères interdits : \\ / ? * [ ] :"); return; }
    if (newName.length > 31) { toast("31 caractères maximum."); return; }
    await run(async (ctx) => {
      ctx.workbook.worksheets.getItem(id).name = newName;
      await ctx.sync();
    }).catch(() => {});
  }

  async function duplicateSheet(id) {
    await run(async (ctx) => {
      const src = ctx.workbook.worksheets.getItem(id);
      const copy = src.copy("After", src);
      copy.activate();
      await ctx.sync();
    }).catch(() => {});
    refresh(true);
  }

  async function toggleHide(id) {
    const s = byId(id);
    if (!s) return;
    const show = s.visibility !== "Visible";
    if (!show && sheets.filter((x) => x.visibility === "Visible").length <= 1) {
      toast("Impossible de masquer la dernière feuille visible.");
      return;
    }
    await run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(id);
      if (!show && id === activeId) {
        const vis = sheets.filter((x) => x.visibility === "Visible" && x.id !== id);
        const next = vis.find((x) => x.position > s.position) || vis[vis.length - 1];
        if (next) ctx.workbook.worksheets.getItem(next.id).activate();
      }
      ws.visibility = show ? "Visible" : "Hidden";
      await ctx.sync();
    }).catch(() => {});
    refresh(true);
  }

  async function setColor(id, color) {
    await run(async (ctx) => {
      ctx.workbook.worksheets.getItem(id).tabColor = color;
      await ctx.sync();
    }).catch(() => {});
    refresh(true);
  }

  async function moveSheet(id, targetId, after) {
    const a = byId(id), b = byId(targetId);
    if (!a || !b || a.id === b.id) return;
    let pos;
    if (after) pos = a.position < b.position ? b.position : b.position + 1;
    else pos = a.position < b.position ? b.position - 1 : b.position;
    if (pos === a.position) return;
    await run(async (ctx) => {
      ctx.workbook.worksheets.getItem(id).position = pos;
      await ctx.sync();
    }).catch(() => {});
    refresh(true);
  }

  function startRename(id) {
    const li = listEl.querySelector('.item[data-id="' + CSS.escape(id) + '"]');
    const s = byId(id);
    if (!li || !s) return;
    busy = true;
    li.draggable = false;
    const nameEl = li.querySelector(".name");
    const input = document.createElement("input");
    input.className = "rename";
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
      refresh(true);
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); finish(true); }
      else if (e.key === "Escape") { e.preventDefault(); finish(false); }
      e.stopPropagation();
    });
    input.addEventListener("blur", () => finish(true));
    input.addEventListener("click", (e) => e.stopPropagation());
    input.addEventListener("dblclick", (e) => e.stopPropagation());
  }

  function openMenu(id, x, y) {
    const s = byId(id);
    if (!s) return;
    menuSheetId = id;
    menuEl.querySelector('[data-act="toggleHide"]').textContent = s.visibility === "Visible" ? "Masquer" : "Afficher";
    menuEl.hidden = false;
    const r = menuEl.getBoundingClientRect();
    menuEl.style.left = Math.max(4, Math.min(x, window.innerWidth - r.width - 4)) + "px";
    menuEl.style.top = Math.max(4, Math.min(y, window.innerHeight - r.height - 4)) + "px";
  }
  function closeMenu() { menuEl.hidden = true; menuSheetId = null; }

  function bindUi() {
    listEl.addEventListener("click", (e) => {
      const li = e.target.closest(".item");
      if (li && !busy) activate(li.dataset.id);
    });
    listEl.addEventListener("dblclick", (e) => {
      const li = e.target.closest(".item");
      if (li) startRename(li.dataset.id);
    });
    listEl.addEventListener("contextmenu", (e) => {
      const li = e.target.closest(".item");
      if (!li) return;
      e.preventDefault();
      openMenu(li.dataset.id, e.clientX, e.clientY);
    });

    menuEl.addEventListener("click", (e) => {
      const btn = e.target.closest("button");
      if (!btn || !menuSheetId) return;
      const id = menuSheetId;
      closeMenu();
      if (btn.dataset.color !== undefined) return setColor(id, btn.dataset.color);
      const act = btn.dataset.act;
      if (act === "rename") startRename(id);
      else if (act === "duplicate") duplicateSheet(id);
      else if (act === "toggleHide") toggleHide(id);
    });
    document.addEventListener("mousedown", (e) => { if (!menuEl.hidden && !menuEl.contains(e.target)) closeMenu(); });
    document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeMenu(); });
    window.addEventListener("blur", closeMenu);

    filterEl.addEventListener("input", render);
    filterEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        const first = listEl.querySelector(".item");
        if (first) activate(first.dataset.id);
      } else if (e.key === "Escape") {
        filterEl.value = ""; render();
      }
    });

    showHiddenEl.checked = store(PREF_HIDDEN) === "1";
    showHiddenEl.addEventListener("change", () => { store(PREF_HIDDEN, showHiddenEl.checked ? "1" : "0"); render(); });

    addBtn.addEventListener("click", addSheet);

    listEl.addEventListener("dragstart", (e) => {
      const li = e.target.closest(".item");
      if (!li) return;
      dragId = li.dataset.id;
      busy = true;
      li.classList.add("dragging");
      e.dataTransfer.effectAllowed = "move";
      try { e.dataTransfer.setData("text/plain", dragId); } catch (_) {}
    });
    listEl.addEventListener("dragover", (e) => {
      if (!dragId) return;
      const li = e.target.closest(".item");
      clearDropMarks();
      if (!li || li.dataset.id === dragId) return;
      e.preventDefault();
      const r = li.getBoundingClientRect();
      li.classList.add(e.clientY > r.top + r.height / 2 ? "drop-after" : "drop-before");
    });
    listEl.addEventListener("drop", (e) => {
      if (!dragId) return;
      e.preventDefault();
      const li = e.target.closest(".item");
      const id = dragId;
      const after = li && li.classList.contains("drop-after");
      endDrag();
      if (li && li.dataset.id !== id) moveSheet(id, li.dataset.id, after);
    });
    listEl.addEventListener("dragend", endDrag);
  }

  function clearDropMarks() {
    listEl.querySelectorAll(".drop-before,.drop-after").forEach((el) => el.classList.remove("drop-before", "drop-after"));
  }
  function endDrag() {
    clearDropMarks();
    listEl.querySelectorAll(".dragging").forEach((el) => el.classList.remove("dragging"));
    dragId = null;
    busy = false;
  }

  async function bindExcelEvents() {
    const handler = () => refresh();
    try {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets;
        const tryOn = (name) => { try { if (ws[name]) ws[name].add(handler); } catch (_) {} };
        ["onActivated", "onAdded", "onDeleted", "onNameChanged", "onVisibilityChanged", "onMoved"].forEach(tryOn);
        await ctx.sync();
      });
    } catch (_) {}
  }

  function startPolling() {
    clearInterval(pollTimer);
    pollTimer = setInterval(() => { if (!document.hidden) refresh(); }, 1500);
  }

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) {
      listEl.innerHTML = '<li class="empty">Ce complément fonctionne dans Excel.</li>';
      return;
    }
    bindUi();
    refresh(true);
    bindExcelEvents();
    startPolling();
  });
})();
