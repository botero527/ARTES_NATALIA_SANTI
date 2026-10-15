/* Núcleo de la UI: puente con Python, i18n, íconos, toasts, modales, tablas. */
"use strict";

/* ── Puente con Python ─────────────────────────────────────────────────── */
async function api(nombre, ...args) {
  try {
    const r = await window.pywebview.api[nombre](...args);
    return r || { ok: false, error: "Sin respuesta" };
  } catch (e) {
    console.error(nombre, e);
    return { ok: false, error: String(e && e.message || e) };
  }
}

/* ── Utilidades ────────────────────────────────────────────────────────── */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }
function el(html) { const t = document.createElement("template"); t.innerHTML = html.trim(); return t.content.firstElementChild; }

/* ── Preferencias (archivo en %LOCALAPPDATA%\AGP_Artes vía Python) ─────── */
const Prefs = {
  _d: {},
  cargar(d) { this._d = d || {}; },
  get(k, def) { return this._d[k] ?? def; },
  set(k, v) { this._d[k] = v; api("guardar_pref", k, v); },
};

/* ── i18n ──────────────────────────────────────────────────────────────── */
let LANG = "es";
function t(key, vars) {
  let s = (I18N[LANG] && I18N[LANG][key]) ?? I18N.es[key] ?? key;
  if (vars) for (const k in vars) s = s.replaceAll(`{${k}}`, vars[k]);
  return s;
}
function aplicarI18n(root = document) {
  $$("[data-i18n]", root).forEach((n) => { n.textContent = t(n.dataset.i18n); });
  $$("[data-i18n-ph]", root).forEach((n) => { n.placeholder = t(n.dataset.i18nPh); });
  $$("[data-lang]").forEach((b) => b.classList.toggle("on", b.dataset.lang === LANG));
  document.documentElement.lang = LANG === "pt" ? "pt-BR" : "es";
}
function setLang(l) {
  LANG = l;
  Prefs.set("lang", l);
  aplicarI18n();
  document.dispatchEvent(new CustomEvent("agp:lang"));
}

/* ── Tema ──────────────────────────────────────────────────────────────── */
function temaActual() { return document.documentElement.dataset.theme === "light" ? "light" : "dark"; }
function setTema(tema, guardar = true) {
  document.documentElement.dataset.theme = tema;
  if (guardar && Prefs.get("tema") !== tema) Prefs.set("tema", tema);
  $$(".js-tema").forEach((b) => {
    b.innerHTML = icon(tema === "dark" ? "sun" : "moon");
    b.setAttribute("aria-label", t(tema === "dark" ? "tema.claro" : "tema.oscuro"));
    b.title = b.getAttribute("aria-label");
  });
  if (window.pywebview && window.pywebview.api) api("set_tema", tema === "dark");
}

/* ── Íconos (SVG lineal estilo Lucide, inline) ─────────────────────────── */
const ICONS = {
  palette: '<circle cx="13.5" cy="6.5" r="1.5"/><circle cx="17.5" cy="10.5" r="1.5"/><circle cx="8.5" cy="7.5" r="1.5"/><circle cx="6.5" cy="12.5" r="1.5"/><path d="M12 2a10 10 0 0 0 0 20c1.1 0 2-.9 2-2v-.5c0-.6.4-1 1-1h1.5A5.5 5.5 0 0 0 22 13 11 10 0 0 0 12 2z"/>',
  database: '<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5"/><path d="M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/>',
  edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>',
  scan: '<path d="M3 7V5a2 2 0 0 1 2-2h2"/><path d="M17 3h2a2 2 0 0 1 2 2v2"/><path d="M21 17v2a2 2 0 0 1-2 2h-2"/><path d="M7 21H5a2 2 0 0 1-2-2v-2"/><path d="M7 8v8M10 8v8M14 8v8M17 8v8"/>',
  users: '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.9"/><path d="M16 3.1a4 4 0 0 1 0 7.8"/>',
  laser: '<path d="M12 2v6"/><path d="M12 16v6"/><path d="M4.9 4.9l4.2 4.2"/><path d="M14.9 14.9l4.2 4.2"/><path d="M2 12h6"/><path d="M16 12h6"/><circle cx="12" cy="12" r="2.5"/>',
  sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
  moon: '<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/>',
  folder: '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>',
  file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/>',
  download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/>',
  zap: '<path d="M13 2L3 14h9l-1 8 10-12h-9z"/>',
  copy: '<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
  trash: '<path d="M3 6h18"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/>',
  refresh: '<path d="M21 12a9 9 0 1 1-2.6-6.4"/><path d="M21 3v6h-6"/>',
  lock: '<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  save: '<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><path d="M17 21v-8H7v8M7 3v5h8"/>',
  x: '<path d="M18 6L6 18M6 6l12 12"/>',
  check: '<path d="M20 6L9 17l-5-5"/>',
  alert: '<path d="M10.3 3.9L1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
  info: '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/>',
  pin: '<path d="M12 17v5"/><path d="M5 17h14l-1.5-2.5V9a4 4 0 0 0-1-2.6L15 5V2H9v3L7.5 6.4a4 4 0 0 0-1 2.6v5.5z"/>',
  pinoff: '<path d="M12 17v5"/><path d="M9 2h6v3l1.5 1.4a4 4 0 0 1 1 2.6v5.5L19 17H5l1.5-2.5V9"/><path d="M2 2l20 20"/>',
  layers: '<path d="M12 2L2 7l10 5 10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/>',
  car: '<path d="M5 17h14v-5l-2-5H7l-2 5z"/><circle cx="7.5" cy="17.5" r="1.5"/><circle cx="16.5" cy="17.5" r="1.5"/>',
  grid: '<rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/>',
  box: '<path d="M21 16V8a2 2 0 0 0-1-1.7l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.7l7 4a2 2 0 0 0 2 0l7-4a2 2 0 0 0 1-1.7z"/>',
  coin: '<circle cx="12" cy="12" r="9"/><path d="M14.5 9a2.5 2.5 0 0 0-2.5-1.5c-1.4 0-2.5.8-2.5 2s1.1 1.7 2.5 2 2.5.8 2.5 2-1.1 2-2.5 2A2.5 2.5 0 0 1 9.5 15M12 6v1.5M12 16.5V18"/>',
  droplet: '<path d="M12 2.7l5.7 5.7a8 8 0 1 1-11.3 0z"/>',
  diamond: '<path d="M12 2l10 10-10 10L2 12z"/>',
  circle: '<circle cx="12" cy="12" r="8"/>',
  target: '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
  external: '<path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><path d="M15 3h6v6M10 14L21 3"/>',
  logout: '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="M16 17l5-5-5-5M21 12H9"/>',
};
function icon(nombre, extra = "") {
  return `<svg class="ic ${extra}" viewBox="0 0 24 24" aria-hidden="true">${ICONS[nombre] || ""}</svg>`;
}
function pintarIconos(root = document) {
  $$("[data-icon]", root).forEach((n) => { n.innerHTML = icon(n.dataset.icon); n.removeAttribute("data-icon"); });
}

/* ── Portapapeles ──────────────────────────────────────────────────────── */
async function copiar(texto, aviso) {
  if (texto == null || texto === "") return;
  const r = await api("copiar", String(texto));
  if (!r.ok) { try { await navigator.clipboard.writeText(String(texto)); } catch (e) {} }
  toast(aviso || `${t("comun.copiado")}: ${String(texto).slice(0, 70)}`, "ok");
}

/* ── Toasts ────────────────────────────────────────────────────────────── */
function toast(msg, tipo = "info", ms = 2600) {
  const ic = { ok: "check", err: "alert", warn: "alert", info: "info" }[tipo] || "info";
  const n = el(`<div class="toast ${tipo}">${icon(ic)}<span>${esc(msg)}</span></div>`);
  $("#toasts").appendChild(n);
  setTimeout(() => { n.style.opacity = "0"; n.style.transition = "opacity .3s"; }, ms);
  setTimeout(() => n.remove(), ms + 350);
}

/* ── Modales ───────────────────────────────────────────────────────────── */
const _pilaModales = [];
/* abrirModal({titulo, sub, icono, color, ancho, cuerpo(html|nodo), pie(html|nodo), alCerrar})
 * alCerrar() se llama al cerrar con X / Esc / clic fuera; si devuelve false no cierra. */
function abrirModal(o) {
  const back = el(`<div class="modal-back" role="dialog" aria-modal="true">
    <div class="modal" style="--mw:${o.ancho || 620}px;--mc:${o.color || "var(--brand)"}">
      <div class="modal-h">
        ${o.icono ? `<div class="mi">${icon(o.icono)}</div>` : ""}
        <div class="grow"><h3>${esc(o.titulo || "")}</h3>${o.sub ? `<p>${esc(o.sub)}</p>` : ""}</div>
        <button class="btn btn-ghost btn-sm btn-icon x" aria-label="${esc(t("comun.cerrar"))}">${icon("x")}</button>
      </div>
      <div class="modal-b"></div>
      <div class="modal-f"></div>
    </div></div>`);
  const cuerpo = $(".modal-b", back), pie = $(".modal-f", back);
  if (typeof o.cuerpo === "string") cuerpo.innerHTML = o.cuerpo; else if (o.cuerpo) cuerpo.appendChild(o.cuerpo);
  if (typeof o.pie === "string") pie.innerHTML = o.pie; else if (o.pie) pie.appendChild(o.pie);
  if (!o.pie) pie.remove();
  const m = {
    back, cuerpo, pie,
    sub: (txt) => { const p = $(".modal-h p", back); if (p) p.textContent = txt; },
    cerrar: () => { back.remove(); const i = _pilaModales.indexOf(m); if (i >= 0) _pilaModales.splice(i, 1); },
    intentarCerrar: async () => { if (o.alCerrar && (await o.alCerrar()) === false) return; m.cerrar(); },
  };
  $(".x", back).addEventListener("click", m.intentarCerrar);
  back.addEventListener("mousedown", (e) => { if (e.target === back && o.cerrarFuera !== false) m.intentarCerrar(); });
  $("#modales").appendChild(back);
  _pilaModales.push(m);
  pintarIconos(back);
  setTimeout(() => { const f = $("[autofocus], input:not([readonly]), select, textarea", cuerpo); if (f) f.focus(); }, 30);
  return m;
}
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && _pilaModales.length) { e.preventDefault(); _pilaModales[_pilaModales.length - 1].intentarCerrar(); }
});
function confirmar(titulo, texto, { ok = "OK", peligro = false } = {}) {
  return new Promise((res) => {
    const m = abrirModal({
      titulo, icono: "alert", color: peligro ? "var(--err)" : "var(--warn)", ancho: 500,
      cuerpo: `<div class="confirm-txt">${esc(texto)}</div>`,
      pie: `<span class="sp"></span><button class="btn" data-r="0">${esc(t("comun.cancelar"))}</button>
            <button class="btn ${peligro ? "btn-danger" : "btn-primary"}" data-r="1">${esc(ok)}</button>`,
      alCerrar: () => { res(false); },
    });
    $$("[data-r]", m.pie).forEach((b) => b.addEventListener("click", () => { res(b.dataset.r === "1"); m.cerrar(); }));
    setTimeout(() => $('[data-r="1"]', m.pie).focus(), 40);
  });
}
function alerta(nivel, titulo, mensaje) {
  const m = abrirModal({
    titulo, icono: "alert", color: nivel === "error" ? "var(--err)" : "var(--warn)", ancho: 520,
    cuerpo: `<div class="confirm-txt">${esc(mensaje)}</div>`,
    pie: `<span class="sp"></span><button class="btn btn-primary">${esc(t("comun.cerrar"))}</button>`,
  });
  $("button.btn-primary", m.pie).addEventListener("click", m.cerrar);
}

/* ── Estados de tabla ──────────────────────────────────────────────────── */
function estadoCarga(cont, filas = 8) {
  cont.innerHTML = Array.from({ length: filas }, () => `<div class="skeleton sk-row"></div>`).join("");
}
function estadoVacio(cont, titulo, detalle) {
  cont.innerHTML = `<div class="state">${icon("search")}<b>${esc(titulo)}</b>${detalle ? `<span>${esc(detalle)}</span>` : ""}</div>`;
}
function estadoError(cont, detalle, reintentar) {
  cont.innerHTML = `<div class="state error">${icon("alert")}<b>${esc(t("comun.error_carga"))}</b>
    <span class="dim" style="max-width:640px;overflow-wrap:anywhere">${esc(String(detalle || "").slice(0, 300))}</span>
    <button class="btn btn-sm">${icon("refresh")}${esc(t("comun.reintentar"))}</button></div>`;
  $("button", cont).addEventListener("click", reintentar);
}

/* ── Tabla con selección de celda y copiado tipo Excel ─────────────────── */
/* o = {headers, filas: [[...]], claves:Set(índices de columna código),
 *      rutas:Set(índices ruta), render(i, j, valor) → html opcional,
 *      alDobleClic(fila, i), alEnter(fila, i), alMenu(fila, i, e)} */
function pintarTabla(cont, o) {
  const claves = o.claves || new Set([0]), rutas = o.rutas || new Set();
  const thead = o.headers.map((h) => `<th>${esc(h)}</th>`).join("");
  const filas = o.filas.map((f, i) => `<tr data-i="${i}">${f.map((v, j) => {
    const vacio = v === "" || v == null;
    const cls = [claves.has(j) ? "k" : "", rutas.has(j) ? "ruta" : "", vacio ? "empty" : ""].join(" ").trim();
    const html = o.render ? o.render(i, j, v) : null;
    return `<td data-j="${j}" class="${cls}" title="${esc(v)}">${html ?? (vacio ? "—" : esc(v))}</td>`;
  }).join("")}</tr>`).join("");
  cont.innerHTML = `<table class="tabla"><thead><tr>${thead}</tr></thead><tbody>${filas}</tbody></table>`;
  cont.tabIndex = 0;
  let sel = { i: -1, j: -1 };

  const marcar = (i, j, scroll) => {
    $$(".cell-sel", cont).forEach((n) => n.classList.remove("cell-sel"));
    $$("tr.sel", cont).forEach((n) => n.classList.remove("sel"));
    sel = { i, j };
    const tr = $(`tr[data-i="${i}"]`, cont);
    if (!tr) return;
    tr.classList.add("sel");
    const td = $(`td[data-j="${j}"]`, tr);
    if (td) { td.classList.add("cell-sel"); if (scroll) td.scrollIntoView({ block: "nearest", inline: "nearest" }); }
  };
  cont.onmousedown = (e) => {
    const td = e.target.closest("td"); if (!td) return;
    marcar(+td.parentElement.dataset.i, +td.dataset.j);
  };
  cont.ondblclick = (e) => {
    const tr = e.target.closest("tr[data-i]"); if (!tr) return;
    const i = +tr.dataset.i;
    if (o.alDobleClic) o.alDobleClic(o.filas[i], i);
    else copiar(o.filas[i].join("\t"), t("comun.fila_copiada"));
  };
  cont.oncontextmenu = (e) => {
    const tr = e.target.closest("tr[data-i]"); if (!tr || !o.alMenu) return;
    e.preventDefault(); o.alMenu(o.filas[+tr.dataset.i], +tr.dataset.i, e);
  };
  cont.onkeydown = (e) => {
    if (sel.i < 0) return;
    const nf = o.filas.length, nc = o.headers.length;
    if ((e.ctrlKey || e.metaKey) && (e.key === "c" || e.key === "C")) {
      e.preventDefault();
      const v = o.filas[sel.i][sel.j];
      if (v !== "" && v != null) copiar(v); else copiar(o.filas[sel.i].join("\t"), t("comun.fila_copiada"));
      return;
    }
    const mov = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] }[e.key];
    if (mov) {
      e.preventDefault();
      marcar(Math.max(0, Math.min(nf - 1, sel.i + mov[0])), Math.max(0, Math.min(nc - 1, sel.j + mov[1])), true);
    } else if (e.key === "Enter" && o.alEnter) {
      e.preventDefault(); o.alEnter(o.filas[sel.i], sel.i);
    }
  };
}

/* ── Consola de log ────────────────────────────────────────────────────── */
function escribirLog(cont, entrada) {
  const vacio = $(".console-empty", cont); if (vacio) vacio.remove();
  const abajo = cont.scrollHeight - cont.scrollTop - cont.clientHeight < 40;
  const n = document.createElement("span");
  n.className = "l";
  n.innerHTML = `<span class="ts">${esc(entrada.ts)}</span><span class="${esc(entrada.tag || "")}">${esc(entrada.msg)}</span>`;
  cont.appendChild(n);
  while (cont.childElementCount > 3000) cont.firstElementChild.remove();
  if (abajo) cont.scrollTop = cont.scrollHeight;
}

/* ── Bus de eventos que llegan del polling ─────────────────────────────── */
const Bus = {
  _h: {},
  on(tipo, fn) { (this._h[tipo] = this._h[tipo] || []).push(fn); },
  emit(tipo, data) { (this._h[tipo] || []).forEach((fn) => { try { fn(data); } catch (e) { console.error(e); } }); },
};
let _ultimoBusy = {};
async function pollLoop() {
  const r = await api("poll");
  if (r.ok) {
    r.logs.forEach((l) => Bus.emit("log:" + l.canal, l));
    r.eventos.forEach((ev) => Bus.emit(ev.tipo, ev));
    for (const k in r.busy) {
      if (r.busy[k] !== _ultimoBusy[k]) Bus.emit("busy:" + k, r.busy[k]);
    }
    _ultimoBusy = r.busy;
  }
  setTimeout(pollLoop, r.logs && r.logs.length >= 400 ? 30 : 300);
}
Bus.on("alerta", (ev) => alerta(ev.nivel, ev.titulo, ev.mensaje));

/* ── Selector numérico (+/−) ───────────────────────────────────────────── */
function numInput(valor, min = 0, max = 20) {
  const n = el(`<div class="num"><button type="button" aria-label="-">−</button><input inputmode="numeric" value="${valor}"/><button type="button" aria-label="+">+</button></div>`);
  const inp = $("input", n);
  const fix = (v) => { v = parseInt(v, 10); if (isNaN(v)) v = min; inp.value = Math.max(min, Math.min(max, v)); };
  n.firstElementChild.onclick = () => fix(+inp.value - 1);
  n.lastElementChild.onclick = () => fix(+inp.value + 1);
  inp.addEventListener("change", () => fix(inp.value));
  n.valor = () => parseInt(inp.value, 10) || 0;
  n.bloquear = (b) => n.classList.toggle("disabled", b);
  return n;
}
