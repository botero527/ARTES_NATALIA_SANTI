/* Pestaña Consultar BD: tarjetas, buscador por tabla y sincronización Excel. */
"use strict";

const BD = (() => {
  const TABS = [
    ["vitrojet", "Vitrojet", "droplet"], ["grandes", "Mallas G", "diamond"], ["pequenas", "Mallas P", "circle"],
    ["vinilos", "Vinilos", "layers"], ["pasta", "Pasta Plata", "coin"], ["glassjet_viejo", "Glassjet V.", "box"],
  ];
  const STATS = [
    ["mallas_grandes", "Mallas G", "diamond", "#3B82F6", "grandes"],
    ["mallas_pequenas", "Mallas P", "circle", "#8B5CF6", "pequenas"],
    ["vitrojet", "Vitrojet", "droplet", "#0ea5e9", "vitrojet"],
    ["pasta_plata", "Pasta Plata", "coin", "#f59e0b", "pasta"],
    ["glassjet_viejo", "Glassjet V.", "box", "#a78bfa", "glassjet_viejo"],
    ["vinilos", "Vinilos", "layers", "#ec4899", "vinilos"],
  ];
  const RUTAS = new Set(["Ruta", "Ruta Archivo"]);
  let tab = "vitrojet", reqId = 0;

  function pintarTabs() {
    $("#bd-tabs").innerHTML = TABS.map(([k, l, ic]) =>
      `<button role="tab" data-k="${k}" class="${k === tab ? "on" : ""}">${icon(ic)} ${esc(l)}</button>`).join("");
    $$("#bd-tabs button").forEach((b) => b.addEventListener("click", () => setTab(b.dataset.k)));
    $$(".stat").forEach((s) => s.classList.toggle("on", s.dataset.tab === tab));
  }
  function setTab(k) { tab = k; pintarTabs(); buscar(); }

  async function buscar() {
    const id = ++reqId;
    const cont = $("#bd-tabla");
    estadoCarga(cont);
    $("#bd-count").textContent = t("comun.cargando");
    const r = await api("bd_consultar", tab, $("#bd-q").value.trim());
    if (id !== reqId) return; // llegó una búsqueda más nueva
    if (!r.ok) {
      estadoError(cont, r.error, buscar);
      $("#bd-count").textContent = t("bd.err");
      return;
    }
    if (!r.rows.length) estadoVacio(cont, t("comun.sin_datos"), t("comun.sin_datos_d"));
    else pintarTabla(cont, {
      headers: r.headers, filas: r.rows,
      rutas: new Set(r.headers.map((h, j) => (RUTAS.has(h) ? j : -1)).filter((j) => j >= 0)),
      alMenu: (fila) => { const j = r.headers.findIndex((h) => RUTAS.has(h)); if (j >= 0 && fila[j]) api("abrir_ruta", fila[j]); },
    });
    $("#bd-count").textContent = t("comun.resultados", { n: r.rows.length, m: r.limite });
  }

  function pintarStats() {
    $("#bd-stats").innerHTML = STATS.map(([k, l, ic, c, tb]) =>
      `<button class="stat" style="--c:${c}" data-k="${k}" data-tab="${tb}">
        <span class="lbl">${icon(ic)}${esc(l)}</span><span class="val skeleton loading"></span></button>`).join("");
    $$(".stat").forEach((s) => s.addEventListener("click", () => setTab(s.dataset.tab)));
  }
  async function cargarStats() {
    const r = await api("bd_stats");
    STATS.forEach(([k]) => {
      const v = $(`.stat[data-k="${k}"] .val`);
      if (!v) return;
      v.classList.remove("skeleton", "loading");
      v.textContent = r.ok ? Number(r.stats[k]).toLocaleString("es-CO") : "—";
    });
  }

  async function sync() {
    const r = await api("sync_excel");
    if (!r.ok) return toast(r.error, "warn");
    $("#bd-sync-log").innerHTML = "";
    $("#bd-sync-det").open = true;
  }
  function syncBusy(on) {
    $("#bd-sync").disabled = on;
    $("#bd-sync-txt").textContent = t(on ? "bd.sincronizando" : "bd.sync");
    $("#bd-sync-prog").classList.toggle("on", on);
    if (on) { const st = $("#bd-sync-st"); st.className = "sync-status"; st.textContent = t("bd.sincronizando"); }
  }

  function refresh() { cargarStats(); buscar(); }

  function init() {
    pintarStats(); pintarTabs();
    $("#bd-q").addEventListener("input", debounce(buscar, 300));
    $("#bd-sync").addEventListener("click", sync);
    Bus.on("log:sync", (l) => escribirLog($("#bd-sync-log"), l));
    Bus.on("busy:sync", syncBusy);
    Bus.on("sync_done", (ev) => {
      const st = $("#bd-sync-st");
      st.textContent = t(ev.exito ? "bd.sync_ok" : "bd.sync_err");
      st.className = "sync-status " + (ev.exito ? "ok" : "warn");
      toast(st.textContent, ev.exito ? "ok" : "warn");
      refresh();
    });
  }
  return { init, refresh };
})();
