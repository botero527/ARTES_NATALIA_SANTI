/* Pestaña Logos Laser: grilla tipo Excel sobre mallas.LOGOS_LASER.
 * Réplica exacta de las reglas de excel_grid.py + tab_logos_laser.py:
 *  - clic en un check lo cambia de una; doble clic / F2 / tecleo editan texto
 *  - Enter baja, Tab / Shift+Tab avanzan y saltan de fila en los bordes
 *  - Supr / Retroceso: check → False, texto → ""
 *  - Ctrl+C: TSV ("1"/"0" en checks); Ctrl+V pega desde la celda ACTIVA
 *  - no se guarda si el valor no cambió (misma regla que _commit_cell)
 *  - el buscador filtra por cliente / producto / id */
"use strict";

/* Reglas puras (las prueba tests/test_equivalencia.py contra excel_grid.py) */
const ReglasGrid = {
  VERDADEROS: new Set(["1", "true", "si", "sí", "x", "yes", "on", "✓"]),
  // Igualdad estilo Python: 1 == True, 0 == False, None != False
  igual(a, b) {
    const num = (v) => typeof v === "boolean" || typeof v === "number";
    if (num(a) && num(b)) return Number(a) === Number(b);
    return a === b;
  },
  vacio: (v) => v === null || v === undefined || v === "",
  sinCambio(viejo, nuevo) { return this.igual(viejo, nuevo) || (this.vacio(viejo) && this.vacio(nuevo)); },
  copiar(filas, cols, r0, r1, c0, c1) {
    const lineas = [];
    for (let r = r0; r <= r1; r++) {
      const vals = [];
      for (let c = c0; c <= c1; c++) {
        const v = filas[r][cols[c].key];
        vals.push(cols[c].type === "bool" ? (v ? "1" : "0") : (v === null || v === undefined ? "" : String(v)));
      }
      lineas.push(vals.join("\t"));
    }
    return lineas.join("\n");
  },
  /* → lista de [r, c, valorNuevo] en el orden en que excel_grid los aplicaría */
  pegar(texto, nFilas, cols, r0, c0) {
    const filas = texto.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    if (filas.length && filas[filas.length - 1] === "") filas.pop();
    const out = [];
    for (let i = 0; i < filas.length; i++) {
      const r = r0 + i;
      if (r >= nFilas) break;
      const vals = filas[i].split("\t");
      for (let j = 0; j < vals.length; j++) {
        const c = c0 + j;
        if (c >= cols.length) break;
        if (cols[c].editable === false) continue;
        const v = vals[j].trim();
        out.push([r, c, cols[c].type === "bool" ? this.VERDADEROS.has(v.toLowerCase()) : v]);
      }
    }
    return out;
  },
  avanzar(r, c, paso, nFilas, nCols) {
    c += paso;
    if (c >= nCols) { c = 0; r = Math.min(r + 1, nFilas - 1); }
    else if (c < 0) { c = nCols - 1; r = Math.max(r - 1, 0); }
    return [r, c];
  },
  filtrar(filas, texto) {
    texto = (texto || "").trim().toLowerCase();
    if (!texto) return filas.slice();
    return filas.filter((f) => ["cliente", "producto", "id"].some((k) => (f[k] || "").toLowerCase().includes(texto)));
  },
};

const Logos = (() => {
  const COLS = [
    { key: "id", header: "ID", width: 100, type: "text", editable: false },
    { key: "cliente", header: "Cliente", width: 190, type: "text" },
    { key: "producto", header: "Producto", width: 190, type: "text" },
    { key: "dxf", header: "DXF", width: 70, type: "bool" },
    { key: "sjf", header: "SJF", width: 70, type: "bool" },
    { key: "validado", header: "Validado", width: 95, type: "bool" },
    { key: "cargado_genesis", header: "Cargado en Genesis", width: 160, type: "bool" },
    { key: "azure", header: "Azure", width: 75, type: "bool" },
    { key: "app", header: "App", width: 70, type: "bool" },
    { key: "observaciones", header: "Observaciones", width: 320, type: "text" },
  ];
  const editable = (c) => COLS[c] && COLS[c].editable !== false;
  const esTexto = (c) => editable(c) && COLS[c].type === "text";
  let todas = [], filas = [];
  let anchor = { r: 0, c: 0 }, active = { r: 0, c: 0 };
  let arrastrando = false, editor = null, cargado = false;
  const wrap = () => $("#ll-grid");

  async function cargar() {
    const w = wrap();
    estadoCarga(w);
    const r = await api("ll_listar");
    if (!r.ok) { estadoError(w, r.error, cargar); $("#ll-status").textContent = ""; return; }
    todas = r.filas; cargado = true;
    filtrar();
  }

  function filtrar() {
    filas = ReglasGrid.filtrar(todas, $("#ll-q").value);
    anchor = { r: 0, c: 0 }; active = { r: 0, c: 0 };
    pintar();
  }

  const celdaHTML = (f, c) => {
    const col = COLS[c], v = f[col.key];
    if (col.type === "bool") return `<span class="xg-chk ${v ? "on" : ""}">${v ? icon("check") : ""}</span>`;
    return esc(v ?? "");
  };

  function pintar() {
    $("#ll-status").textContent = t("ll.filas", { n: filas.length });
    if (!filas.length) { estadoVacio(wrap(), t("comun.sin_datos"), ""); return; }
    const cols = COLS.map((c) => `<col style="width:${c.width}px">`).join("");
    const th = COLS.map((c) => `<th>${esc(c.header)}</th>`).join("");
    const body = filas.map((f, r) => `<tr>${COLS.map((col, c) =>
      `<td data-r="${r}" data-c="${c}" class="${col.type === "bool" ? "c-center" : ""} ${editable(c) ? "" : "ro"}" title="${col.type === "bool" ? "" : esc(f[col.key] ?? "")}">${celdaHTML(f, c)}</td>`).join("")}</tr>`).join("");
    wrap().innerHTML = `<table class="xg"><colgroup>${cols}</colgroup><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table>`;
    marcarSeleccion();
  }

  const td = (r, c) => $(`td[data-r="${r}"][data-c="${c}"]`, wrap());
  function rango() {
    return { r1: Math.min(anchor.r, active.r), r2: Math.max(anchor.r, active.r),
             c1: Math.min(anchor.c, active.c), c2: Math.max(anchor.c, active.c) };
  }
  function marcarSeleccion(scroll) {
    $$("td.in-range, td.active", wrap()).forEach((n) => n.classList.remove("in-range", "active"));
    const g = rango();
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const n = td(r, c); if (n) n.classList.add("in-range"); }
    const a = td(active.r, active.c);
    if (a) { a.classList.add("active"); if (scroll) a.scrollIntoView({ block: "nearest", inline: "nearest" }); }
  }
  function mover(dr, dc, extender) {
    terminar(true);
    if (!filas.length) return;
    active = { r: Math.max(0, Math.min(filas.length - 1, active.r + dr)), c: Math.max(0, Math.min(COLS.length - 1, active.c + dc)) };
    if (!extender) anchor = { ...active };
    marcarSeleccion(true);
    wrap().focus();
  }
  function avanzar(paso) {
    terminar(true);
    if (!filas.length) return;
    const [r, c] = ReglasGrid.avanzar(active.r, active.c, paso, filas.length, COLS.length);
    anchor = active = { r, c };
    marcarSeleccion(true);
    wrap().focus();
  }

  /* ── Guardado: igual que ExcelGrid._commit_cell + TabLogosLaser._on_edit ── */
  async function guardar(r, c, nuevo) {
    const f = filas[r], col = COLS[c];
    if (!f || !editable(c)) return true;
    const viejo = f[col.key];
    if (ReglasGrid.sinCambio(viejo, nuevo)) return true;
    const n = td(r, c);
    if (n) { n.classList.add("saving"); n.classList.remove("fail"); }
    const res = await api("ll_actualizar", f.id, col.key, nuevo);
    const n2 = td(filas.indexOf(f), c);
    if (n2) n2.classList.remove("saving");
    if (!res.ok) {
      if (n2) n2.classList.add("fail");
      toast(t("ll.err_guardar", { e: res.error }), "err", 5000);
      return false;
    }
    f[col.key] = nuevo;
    if (n2) { n2.innerHTML = celdaHTML(f, c); if (col.type !== "bool") n2.title = nuevo ?? ""; }
    return true;
  }

  /* ── Edición de texto (Entry flotante) ───────────────────────────────── */
  function editar(textoInicial, seleccionarTodo) {
    terminar(true);
    const { r, c } = active;
    if (!esTexto(c) || !filas[r]) return;
    const n = td(r, c); if (!n) return;
    const inp = document.createElement("input");
    inp.className = "xg-editor";
    inp.value = textoInicial ?? String(filas[r][COLS[c].key] || "");
    n.appendChild(inp);
    editor = { inp, r, c };
    inp.focus();
    if (seleccionarTodo) inp.select(); else inp.setSelectionRange(inp.value.length, inp.value.length);
    inp.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Enter") { e.preventDefault(); terminar(true); mover(1, 0); }
      else if (e.key === "Tab") { e.preventDefault(); terminar(true); avanzar(e.shiftKey ? -1 : 1); }
      else if (e.key === "Escape") { e.preventDefault(); terminar(false); wrap().focus(); }
    });
    inp.addEventListener("blur", () => terminar(true));
    inp.addEventListener("mousedown", (e) => e.stopPropagation());
  }
  function terminar(commit) {
    if (!editor) return;
    const { inp, r, c } = editor;
    editor = null;
    const valor = inp.value;          // sin recortar: logos_laser.actualizar_campo hace el strip
    inp.remove();
    if (commit) guardar(r, c, valor);
  }

  async function pegar(texto) {
    const cambios = ReglasGrid.pegar(texto, filas.length, COLS, active.r, active.c);
    let n = 0;
    for (const [r, c, v] of cambios) if (await guardar(r, c, v)) n++;
    toast(`✔ ${n}`, "ok", 1500);
  }
  function limpiar() {
    const g = rango();
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++)
      if (editable(c)) guardar(r, c, COLS[c].type === "bool" ? false : "");
  }

  async function nuevaFila() {
    const b = $("#ll-nueva"); b.disabled = true;
    const r = await api("ll_crear");
    b.disabled = false;
    if (!r.ok) return alerta("error", t("ll.nueva"), r.error);
    todas.unshift(r.fila);
    $("#ll-q").value = "";
    filtrar();
    wrap().focus();
    toast(t("ll.creada", { id: r.fila.id }), "ok");
  }

  function init() {
    const w = wrap();
    w.addEventListener("mousedown", (e) => {
      const n = e.target.closest("td[data-r]"); if (!n) return;
      terminar(true);
      const p = { r: +n.dataset.r, c: +n.dataset.c };
      if (e.shiftKey) active = p;
      else {
        if (COLS[p.c].type === "bool" && editable(p.c)) guardar(p.r, p.c, !filas[p.r][COLS[p.c].key]);
        anchor = p; active = p;
      }
      arrastrando = true;
      marcarSeleccion();
      w.focus();
      e.preventDefault();
    });
    w.addEventListener("mouseover", (e) => {
      if (!arrastrando) return;
      const n = e.target.closest("td[data-r]"); if (!n) return;
      active = { r: +n.dataset.r, c: +n.dataset.c }; marcarSeleccion();
    });
    document.addEventListener("mouseup", () => { arrastrando = false; });
    w.addEventListener("dblclick", (e) => { if (e.target.closest("td[data-r]")) editar(null, true); });
    w.addEventListener("keydown", (e) => {
      if (editor || !filas.length) return;
      const k = e.key, ctrl = e.ctrlKey || e.metaKey;
      if (ctrl && (k === "a" || k === "A")) { e.preventDefault(); anchor = { r: 0, c: 0 }; active = { r: filas.length - 1, c: COLS.length - 1 }; marcarSeleccion(); return; }
      if (ctrl || e.altKey) return; // Ctrl+C / Ctrl+V los manejan los eventos copy / paste
      const mov = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] }[k];
      if (mov) { e.preventDefault(); mover(mov[0], mov[1], e.shiftKey); }
      else if (k === "Tab") { e.preventDefault(); avanzar(e.shiftKey ? -1 : 1); }
      else if (k === "Enter") { e.preventDefault(); mover(1, 0); }
      else if (k === "F2") { e.preventDefault(); editar(null, false); }
      else if (k === "Delete" || k === "Backspace") { e.preventDefault(); limpiar(); }
      else if (k.length === 1 && esTexto(active.c)) { e.preventDefault(); editar(k, false); }
    });
    document.addEventListener("copy", (e) => {
      if (editor || document.activeElement !== w || !filas.length) return;
      const g = rango();
      e.preventDefault(); e.clipboardData.setData("text/plain", ReglasGrid.copiar(filas, COLS, g.r1, g.r2, g.c1, g.c2));
      toast(t("comun.copiado"), "ok", 1200);
    });
    document.addEventListener("paste", (e) => {
      if (editor || document.activeElement !== w || !filas.length) return;
      e.preventDefault(); pegar(e.clipboardData.getData("text/plain") || "");
    });
    $("#ll-q").addEventListener("input", debounce(filtrar, 150));
    $("#ll-nueva").addEventListener("click", nuevaFila);
    $("#ll-reload").addEventListener("click", cargar);
  }

  return { init, refresh: cargar, alMostrar: () => { if (!cargado) cargar(); } };
})();
