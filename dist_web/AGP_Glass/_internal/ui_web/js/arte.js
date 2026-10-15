/* Pestaña Crear Arte + cajetín + auto-asignación vitro/malla. */
"use strict";

/* Reglas del cajetín — réplica de parsear_cod / extraer_anio / _sugerir_nombre_arte
 * de crear_arte_acad.dialogo_cajetin (las prueba tests/test_equivalencia.py). */
const ReglasCajetin = {
  parsearCod(cod) {
    const g = (cod.trim().match(/[\d-]+/g) || []).map((x) => x.replace(/^-+|-+$/g, "")).filter(Boolean);
    return [g[0] || "", g[1] || "", g[2] || ""];
  },
  extraerAnio(txt) { const m = txt.match(/\b(19|20)\d{2}\b/); return m ? m[0] : ""; },
  nombreSugerido(malla, pieza, nombrePlano) {
    malla = malla.trim(); pieza = pieza.trim();
    const mallaLimpia = malla ? malla.replace(/,/g, "/").split("/").map((m) => m.trim()).filter(Boolean).join(" - ") : "";
    const partes = [mallaLimpia, pieza].filter(Boolean);
    let s = (partes.length ? "P " + partes.join(" ") : "P " + nombrePlano) + ".dwg";
    s = s.replace(/[/\\:*?"<>|]/g, "-");
    while (s.includes("  ")) s = s.replaceAll("  ", " ");
    while (s.includes("--")) s = s.replaceAll("--", "-");
    return s;
  },
};

const Arte = (() => {
  let resultados = [];
  let lanzado = false;

  const base = () => $("#arte-base").value.trim().replace(/^"|"$/g, "");
  const dwg = () => $("#arte-dwg").value.trim().replace(/^"|"$/g, "");

  function guardarRutas() {
    if (Prefs.get("arte_base") !== base()) Prefs.set("arte_base", base());
    if (Prefs.get("arte_dwg") !== dwg()) Prefs.set("arte_dwg", dwg());
  }

  async function validar({ reqBase, reqDwg }) {
    const r = await api("arte_validar", base(), dwg());
    if (!r.ok) { toast(r.error, "err"); return false; }
    if (reqBase && !r.base_ok) { alerta("warning", t("arte.config"), t("arte.req_base")); $("#arte-base").focus(); return false; }
    if (reqDwg && !r.dwg_ok) { alerta("warning", t("arte.config"), t("arte.req_dwg")); $("#arte-dwg").focus(); return false; }
    return true;
  }

  async function accion(tipo) {
    guardarRutas();
    if (tipo === "extraer") {
      if (!(await validar({ reqBase: true, reqDwg: true }))) return;
      const r = await api("arte_extraer", dwg());
      if (!r.ok) return toast(r.error, "err");
      lanzado = true; toast(t("arte.lanzado"), "info");
    } else if (tipo === "buscar") {
      if (!(await validar({ reqBase: true, reqDwg: false }))) return;
      estadoCarga($("#arte-tabla"), 4);
      const r = await api("arte_buscar", base(), dwg());
      if (!r.ok) { estadoError($("#arte-tabla"), r.error, () => accion("buscar")); return; }
      pintarResultados(r.resultados);
    } else if (tipo === "crear" || tipo === "todo") {
      if (!(await validar({ reqBase: tipo === "todo", reqDwg: true }))) return;
      const nombre = dwg().split("\\").pop().replace(/\.[^.]+$/, "");
      abrirCajetin(nombre, tipo);
    }
  }

  function pintarResultados(res) {
    resultados = res || [];
    $("#arte-count").textContent = `${resultados.length} ${resultados.length === 1 ? "resultado" : "resultados"}`;
    if (!resultados.length) { estadoVacio($("#arte-tabla"), t("comun.sin_datos"), t("arte.sin_busqueda")); return; }
    pintarTabla($("#arte-tabla"), {
      headers: [t("arte.version"), t("arte.archivo"), t("arte.tipo")],
      filas: resultados.map((r) => [r.version, r.archivo, r.archivo.split(".").pop().toUpperCase()]),
      claves: new Set([1]),
      alDobleClic: (_f, i) => superponer(resultados[i]),
      alEnter: (_f, i) => superponer(resultados[i]),
      alMenu: (_f, i) => api("abrir_ruta", resultados[i].ruta_completa),
    });
  }

  async function superponer(r) {
    const res = await api("arte_superponer", r.ruta_completa, dwg());
    if (!res.ok) return alerta("warning", t("arte.buscar"), res.error);
    lanzado = true;
  }

  function busy(on) {
    $$(".wf-btn").forEach((b) => (b.disabled = on));
    $("#arte-prog").classList.toggle("on", on);
    $("#arte-busy").classList.toggle("hide", !on);
    if (!on && lanzado) { lanzado = false; toast(t("arte.terminado"), "ok"); }
  }

  /* ── Cajetín ─────────────────────────────────────────────────────────── */
  const hoy = () => { const d = new Date(); return `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`; };
  const SECCIONES = () => [
    ["caj.datos", [
      ["DIBUJO", "Dibujo", null], ["VEHICULO", "Nombre vehículo", null],
      ["MODELO", "Modelo / Año", "auto"], ["COD PLANO", "Código de plano", null],
      ["NAGS", "NAGS", "auto"], ["VERSION", "Versión", "auto"], ["PIEZA", "Pieza", "auto"],
      ["VITRO", "Vitro", null], ["MALLA", "Malla", null], ["NOMBRE ARTE", "Nombre archivo arte", "auto"],
    ]],
    ["caj.defecto", [
      ["FECHA", "Fecha", hoy()], ["REVISADO", "Revisado por", "Santiago P."],
      ["MEDIDAS", "Medidas", "Milimetros"], ["VISTA", "Vista", "Interna"], ["ESCALA", "Escala", "1:1"],
    ]],
  ];

  const parsearCod = (cod) => ReglasCajetin.parsearCod(cod);
  const extraerAnio = (txt) => ReglasCajetin.extraerAnio(txt);

  function abrirCajetin(nombrePlano, modo) {
    let rid = null; // reserva PENDIENTE en BD (si usó auto-asignar)
    const cuerpo = document.createElement("div");
    SECCIONES().forEach(([sec, campos]) => {
      cuerpo.appendChild(el(`<div class="sec-t">${esc(t(sec))}</div>`));
      const g = el(`<div class="form-grid"></div>`);
      campos.forEach(([k, lbl, def]) => {
        const full = k === "NOMBRE ARTE" ? " full" : "";
        g.appendChild(el(`<div class="field${full}"><label>${esc(lbl)}${def === "auto" ? `<span class="badge-auto">auto</span>` : ""}</label>
          <input class="input" data-k="${esc(k)}" spellcheck="false" value="${def && def !== "auto" ? esc(def) : ""}" /></div>`));
      });
      cuerpo.appendChild(g);
    });
    cuerpo.appendChild(el(`<div class="sec-t">${esc(t("caj.asignacion"))}</div>`));
    const asigRow = el(`<div class="row"><button class="btn btn-primary" type="button">${icon("target")}${esc(t("caj.auto_asignar"))}</button>
      <span class="asig-preview">${esc(t("caj.sin_asignar"))}</span></div>`);
    cuerpo.appendChild(asigRow);

    const E = (k) => $(`input[data-k="${k}"]`, cuerpo);
    const setAuto = (k, v) => { E(k).value = v; E(k).classList.add("auto"); };
    // Igual que on_cod_plano del cajetín Tk: NO recalcula NOMBRE ARTE (solo al abrir y con MALLA/PIEZA)
    const onCod = () => { const [n, v, p] = parsearCod(E("COD PLANO").value); setAuto("NAGS", n); setAuto("VERSION", v); setAuto("PIEZA", p); };
    const onVeh = () => { const a = extraerAnio(E("VEHICULO").value); if (a) setAuto("MODELO", a); };
    function sugerirNombre() {
      setAuto("NOMBRE ARTE", ReglasCajetin.nombreSugerido(E("MALLA").value, E("PIEZA").value, nombrePlano));
    }
    E("COD PLANO").addEventListener("blur", onCod);
    E("VEHICULO").addEventListener("blur", onVeh);
    ["MALLA", "PIEZA"].forEach((k) => { E(k).addEventListener("input", sugerirNombre); E(k).addEventListener("blur", sugerirNombre); });
    $$("input", cuerpo).forEach((inp) => inp.addEventListener("input", () => inp.classList.remove("auto")));
    // Valores por defecto: si lo dejan vacío vuelve el default (igual que el cajetín Tk)
    SECCIONES()[1][1].forEach(([k, , def]) => E(k).addEventListener("blur", () => { if (!E(k).value.trim()) E(k).value = def; }));
    // Enter avanza al siguiente campo
    const inputs = $$("input", cuerpo);
    inputs.forEach((inp, i) => inp.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); const n = inputs[(i + 1) % inputs.length]; n.focus(); n.select(); }
    }));

    if (nombrePlano) { E("COD PLANO").value = nombrePlano; onCod(); sugerirNombre(); }

    const pie = el(`<div style="display:contents">
      <span class="hint">${esc(t("caj.hint"))}</span><span class="sp"></span>
      <button class="btn" data-a="cancel">${icon("x")}${esc(t("comun.cancelar"))}</button>
      <button class="btn btn-primary" data-a="ok">${icon("check")}${esc(t("caj.aplicar"))}</button></div>`);

    const cancelarReserva = async () => { if (rid) { await api("cancelar_reserva", rid); rid = null; } };
    const m = abrirModal({
      titulo: `AGP Arte Maker — ${t("caj.titulo")}`, sub: nombrePlano || "nuevo plano",
      icono: modo === "todo" ? "zap" : "palette",
      color: modo === "todo" ? "var(--c-todo)" : "var(--c-crear)", ancho: 760,
      cuerpo, pie, cerrarFuera: false,
      alCerrar: async () => { await cancelarReserva(); },
    });
    setTimeout(() => E("DIBUJO").focus(), 60);

    $("button", asigRow).addEventListener("click", async () => {
      if (rid) { await cancelarReserva(); E("VITRO").value = ""; E("MALLA").value = ""; }
      const prev = $(".asig-preview", asigRow);
      prev.textContent = t("caj.sin_asignar"); prev.classList.remove("ok");
      const res = await dialogoAsignacion(nombrePlano);
      if (!res) return;
      rid = res.rid;
      const partes = [];
      if (res.vitros.length) { partes.push("Vitros: " + res.vitros.join(", ")); setAuto("VITRO", res.vitros.join(" / ")); }
      const mallas = res.grandes.concat(res.pequenas);
      if (mallas.length) { partes.push("Mallas: " + mallas.join(", ")); setAuto("MALLA", mallas.join(" / ")); sugerirNombre(); }
      prev.textContent = partes.join(" | ") || "(sin asignación)"; prev.classList.add("ok");
    });

    $('[data-a="cancel"]', m.pie).addEventListener("click", m.intentarCerrar);
    $('[data-a="ok"]', m.pie).addEventListener("click", async () => {
      const valores = {};
      inputs.forEach((inp) => (valores[inp.dataset.k] = inp.value.trim()));
      const btn = $('[data-a="ok"]', m.pie); btn.disabled = true;
      const r = await api("arte_iniciar", modo, dwg(), valores, rid, $("#arte-comp").checked);
      btn.disabled = false;
      if (!r.ok) { alerta("warning", t("caj.titulo"), r.error); return; }
      rid = null; // ya quedó confirmada en BD
      m.cerrar();
      lanzado = true;
      toast(t("arte.lanzado"), "info");
    });
  }

  /* Diálogo de cantidad → reserva real (PENDIENTE) → devuelve {rid, vitros, grandes, pequenas} */
  function dialogoAsignacion(nombrePlano) {
    return new Promise((resolver) => {
      let actual = null;
      const nv = numInput(1), ng = numInput(0), np = numInput(0);
      const cuerpo = el(`<div>
        <div class="sec-t">${esc(t("asg.cantidad"))}</div>
        <div class="qty-row"><div><b>${esc(t("asg.vitros"))}</b><small>T-xxxx</small></div><span data-n="v"></span></div>
        <div class="qty-row"><div><b>${esc(t("asg.grandes"))}</b><small>A-xxxx</small></div><span data-n="g"></span></div>
        <div class="qty-row"><div><b>${esc(t("asg.pequenas"))}</b><small>número</small></div><span data-n="p"></span></div>
        <div class="sec-t" style="margin-top:18px">${esc(t("asg.numeros"))}</div>
        <div class="codes"><div class="dim" data-st>${esc(t("asg.info"))}</div><div data-codes></div></div></div>`);
      $('[data-n="v"]', cuerpo).replaceWith(nv); $('[data-n="g"]', cuerpo).replaceWith(ng); $('[data-n="p"]', cuerpo).replaceWith(np);
      const pie = el(`<div style="display:contents"><span class="sp"></span>
        <button class="btn" data-a="cancel">${esc(t("comun.cancelar"))}</button>
        <button class="btn" data-a="res">${icon("lock")}${esc(t("asg.reservar"))}</button>
        <button class="btn btn-primary" data-a="ok">${icon("check")}${esc(t("asg.confirmar"))}</button></div>`);
      let terminado = false;
      const m = abrirModal({
        titulo: t("asg.titulo"), sub: nombrePlano ? `Plano: ${nombrePlano}` : "", icono: "target", ancho: 520,
        cuerpo, pie, cerrarFuera: false,
        alCerrar: async () => { if (!terminado) { if (actual) await api("cancelar_reserva", actual.rid); resolver(null); } },
      });
      const st = $("[data-st]", cuerpo), codes = $("[data-codes]", cuerpo);
      async function reservar() {
        st.textContent = t("asg.reservando"); st.className = "dim"; codes.innerHTML = "";
        const r = await api("reservar", nv.valor(), ng.valor(), np.valor(), actual ? actual.rid : null);
        if (!r.ok) { st.textContent = r.error === "Indica al menos 1 vitro o malla." ? t("asg.min") : `Error: ${r.error}`; st.className = "msg err"; actual = null; return false; }
        actual = r;
        codes.innerHTML = lineasCodigos(r);
        st.textContent = t("asg.pendiente"); st.className = "msg ok";
        if (r.aviso) toast(r.aviso, "warn", 6000);
        enlazarCopiar(codes);
        return true;
      }
      $('[data-a="res"]', pie).addEventListener("click", reservar);
      $('[data-a="cancel"]', pie).addEventListener("click", m.intentarCerrar);
      $('[data-a="ok"]', pie).addEventListener("click", async () => {
        if (!actual && !(await reservar())) return;
        terminado = true; m.cerrar(); resolver(actual);
      });
    });
  }

  function init() {
    $("#arte-base").value = Prefs.get("arte_base", "");
    $("#arte-dwg").value = Prefs.get("arte_dwg", "");
    $$(".wf-btn").forEach((b) => b.addEventListener("click", () => accion(b.dataset.accion)));
    $("#arte-base-pick").addEventListener("click", async () => {
      const r = await api("elegir_carpeta", base());
      if (r.ok && r.ruta) { $("#arte-base").value = r.ruta; guardarRutas(); }
    });
    $("#arte-dwg-pick").addEventListener("click", async () => {
      const r = await api("elegir_archivo", "dwg", base());
      if (r.ok && r.ruta) { $("#arte-dwg").value = r.ruta; guardarRutas(); }
    });
    ["#arte-base", "#arte-dwg"].forEach((s) => $(s).addEventListener("change", () => {
      $(s).value = $(s).value.trim().replace(/^"|"$/g, ""); guardarRutas();
    }));
    $("#arte-log-clear").addEventListener("click", () => { $("#arte-log").innerHTML = ""; });
    $("#arte-log-copy").addEventListener("click", () => copiar($("#arte-log").innerText, t("comun.copiado")));
    Bus.on("log:arte", (l) => escribirLog($("#arte-log"), l));
    Bus.on("busy:arte", busy);
    estadoVacio($("#arte-tabla"), t("comun.sin_datos"), t("arte.sin_busqueda"));
  }

  return { init, pintarResultados: () => pintarResultados(resultados) };
})();

/* Líneas de códigos reservados con botón copiar (cajetín y Separar) */
function lineasCodigos(r) {
  const linea = (lbl, cods) => !cods.length ? "" : `<div class="code-line"><span class="lbl">${esc(lbl)}</span>${cods.map((c) =>
    `<span class="code-chip">${esc(c)}<button type="button" data-copy="${esc(c)}" aria-label="Copiar ${esc(c)}">${icon("copy")}</button></span>`).join("")}</div>`;
  return linea("Vitro", r.vitros) + linea("Malla G", r.grandes) + linea("Malla P", r.pequenas);
}
function enlazarCopiar(root) {
  $$("[data-copy]", root).forEach((b) => b.addEventListener("click", () => copiar(b.dataset.copy)));
}
