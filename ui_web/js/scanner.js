/* Pestaña Scanner: orden (pistola) → ZFER → vitro / mallas / rutas. */
"use strict";

const Scanner = (() => {
  const recientes = [];
  let ocupado = false;

  async function buscar(orden) {
    const inp = $("#sc-input");
    orden = (orden ?? inp.value).trim();
    if (!orden || ocupado) return;
    ocupado = true;
    inp.disabled = true; $("#sc-btn").disabled = true;
    $("#sc-prog").classList.add("on");
    const r = await api("scanner_buscar", orden);
    ocupado = false;
    inp.disabled = false; $("#sc-btn").disabled = false;
    $("#sc-prog").classList.remove("on");
    if (!r.ok) {
      inp.classList.add("bad"); setTimeout(() => inp.classList.remove("bad"), 2000);
      const msg = r.no_encontrada ? t("sc.no_encontrada", { o: orden }) : r.error;
      $("#sc-res").innerHTML = `<div class="card"><div class="state error">${icon("alert")}<b>${esc(msg)}</b>
        <button class="btn btn-sm">${icon("refresh")}${esc(t("comun.reintentar"))}</button></div></div>`;
      $("#sc-res button").addEventListener("click", () => buscar(orden));
      inp.focus(); inp.select();
      return;
    }
    pintar(r);
    if (!recientes.includes(orden)) { recientes.unshift(orden); recientes.length = Math.min(recientes.length, 8); }
    pintarRecientes();
    inp.value = ""; inp.focus();
  }

  function pintarRecientes() {
    const h = $("#sc-hist");
    h.classList.toggle("hide", !recientes.length);
    h.innerHTML = `<span>${esc(t("sc.recientes"))}</span>` + recientes.map((o) => `<button data-o="${esc(o)}">${esc(o)}</button>`).join("");
    $$("button", h).forEach((b) => b.addEventListener("click", () => buscar(b.dataset.o)));
  }

  const btnCopiar = (v) => v ? `<button class="btn btn-ghost btn-sm btn-icon" data-copy="${esc(v)}" aria-label="Copiar">${icon("copy")}</button>` : "";

  function pintar(r) {
    const vitroRuta = r.ruta_err
      ? `<div class="ruta-box" style="color:var(--err-ink)">${esc(t("sc.error_bd", { e: r.ruta_err.slice(0, 80) }))}</div>`
      : `<div class="ruta-box"><span>${esc(r.vitro_ruta || t("sc.sin_ruta"))}</span>${r.vitro_ruta ? btnCopiar(r.vitro_ruta) +
          `<button class="btn btn-ghost btn-sm btn-icon" data-open="${esc(r.vitro_ruta)}" aria-label="${esc(t("comun.abrir_carpeta"))}">${icon("folder")}</button>` : ""}</div>`;
    const mallas = r.mallas.map((m, i) => `<div class="malla">
        <div class="malla-h"><span class="malla-n">${i + 1}</span><b class="mono">${esc(m.codigo)}</b>${btnCopiar(m.codigo.split(" ")[0])}</div>
        ${m.misma_carpeta ? `<div class="misma">${icon("check")}${esc(t("sc.misma"))}</div>`
          : `<div class="ruta-box"><span>${esc(m.ruta || t("sc.sin_ruta"))}</span>${m.ruta ? btnCopiar(m.ruta) +
              `<button class="btn btn-ghost btn-sm btn-icon" data-open="${esc(m.ruta)}" aria-label="${esc(t("comun.abrir_carpeta"))}">${icon("folder")}</button>` : ""}</div>`}
      </div>`).join("");
    $("#sc-res").innerHTML = `
      ${r.alerta ? `<div class="alerta" style="margin-bottom:16px">${icon("alert")}<span>${esc(r.alerta)}</span>
        <button class="btn btn-sm" data-copy="${esc(r.alerta)}">${icon("copy")}${esc(t("sc.copiar_alerta"))}</button></div>` : ""}
      <div class="scan-grid">
        <div class="stack">
          <div class="ids">
            <div class="id-blk orden"><small>${esc(t("sc.orden").toUpperCase())}</small><div class="v">${esc(r.orden)}${btnCopiar(r.orden)}</div></div>
            <div class="id-blk zfer"><small>ZFER</small><div class="v">${esc(r.zfer)}${btnCopiar(r.zfer)}</div></div>
          </div>
          <div class="card acc" style="--c:#58a6ff">
            <div class="veh-top">${icon("car")}<h3>${esc(r.vehiculo)}</h3>
              <span class="logo-chip ${r.logo ? "" : "none"}">${esc(r.logo || "—")}</span></div>
            <div class="kv">
              <div><small>${esc(t("sc.lote").toUpperCase())}</small><b style="color:#e3b341">${esc(r.lote)}</b></div>
              <div><small>${esc(t("sc.version").toUpperCase())}</small><b style="color:#58a6ff">${esc(r.version)}</b></div>
              <div><small>${esc(t("sc.codveh").toUpperCase())}</small><b style="color:#79c0ff">${esc(r.cod_veh)}</b></div>
              <div><small>${esc(t("sc.tipo").toUpperCase())}</small><b style="color:#d2a8ff">${esc(r.tipo_pieza)}</b></div>
              <div><small>${esc(t("sc.traz").toUpperCase())}</small><b style="color:#ffa657">${esc(r.trazabilidad)}</b></div>
            </div>
          </div>
          <div class="card acc" style="--c:var(--ok)">
            <div class="card-h"><h2 style="color:var(--ok-ink)">${esc(t("sc.vitro"))}</h2></div>
            <div class="vitro-big">${esc(r.vitro)}${r.vitro !== "—" ? btnCopiar(r.vitro) : ""}</div>
            ${vitroRuta}
          </div>
        </div>
        <div class="card acc" style="--c:var(--violet)">
          <div class="card-h"><h2 style="color:#b794f6">${esc(t("sc.mallas"))}</h2><span class="count">${r.mallas.length || ""}</span></div>
          <div class="mallas-list">${mallas || `<div class="state">${icon("grid")}<span>—</span></div>`}</div>
        </div>
      </div>`;
    $$("[data-copy]", $("#sc-res")).forEach((b) => b.addEventListener("click", () => copiar(b.dataset.copy)));
    $$("[data-open]", $("#sc-res")).forEach((b) => b.addEventListener("click", async () => {
      const x = await api("abrir_ruta", b.dataset.open); if (!x.ok) toast(x.error, "warn");
    }));
  }

  function alMostrar() { setTimeout(() => $("#sc-input").focus(), 50); }

  function init() {
    $("#sc-input").addEventListener("keydown", (e) => { if (e.key === "Enter") buscar(); });
    $("#sc-btn").addEventListener("click", () => buscar());
    $("#sc-res").innerHTML = `<div class="card"><div class="state">${icon("scan")}<span>${esc(t("sc.inicio"))}</span></div></div>`;
  }
  return { init, alMostrar };
})();
