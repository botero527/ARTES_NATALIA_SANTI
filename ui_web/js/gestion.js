/* Pestaña Gestión BD: editar (doble clic), anular, insertar manual, separar. */
"use strict";

const Gestion = (() => {
  const TABS = [
    ["vitrojet", "Vitrojet", "droplet"], ["grandes", "Mallas G", "diamond"], ["pequenas", "Mallas P", "circle"],
    ["vinilos", "Vinilos", "layers"], ["pasta_plata", "Pasta Plata", "coin"],
  ];
  const RUTAS = new Set(["ruta", "ruta_dwg", "ruta_archivo"]);
  let tab = "vitrojet", reqId = 0, datos = null;

  function pintarTabs() {
    $("#ge-tabs").innerHTML = TABS.map(([k, l, ic]) =>
      `<button role="tab" data-k="${k}" class="${k === tab ? "on" : ""}">${icon(ic)} ${esc(l)}</button>`).join("");
    $$("#ge-tabs button").forEach((b) => b.addEventListener("click", () => { tab = b.dataset.k; pintarTabs(); buscar(); }));
  }

  async function buscar() {
    const id = ++reqId, cont = $("#ge-tabla");
    estadoCarga(cont);
    $("#ge-count").textContent = t("comun.cargando");
    const r = await api("gestion_buscar", tab, $("#ge-q").value.trim());
    if (id !== reqId) return;
    if (!r.ok) { estadoError(cont, r.error, buscar); $("#ge-count").textContent = t("bd.err"); return; }
    datos = r;
    if (!r.rows.length) { estadoVacio(cont, t("comun.sin_datos"), t("comun.sin_datos_d")); }
    else {
      const jEstado = r.fields.indexOf("estado");
      pintarTabla(cont, {
        headers: r.headers,
        filas: r.rows.map((row) => r.fields.map((f) => row[f])),
        rutas: new Set(r.fields.map((f, j) => (RUTAS.has(f) ? j : -1)).filter((j) => j >= 0)),
        render: (_i, j, v) => (j === jEstado ? `<span class="estado ${esc(v || "NA")}">${esc(v || "—")}</span>` : null),
        alDobleClic: (_f, i) => editar(r.rows[i]),
        alEnter: (_f, i) => editar(r.rows[i]),
        alMenu: (_f, i) => { const f = r.fields.find((x) => RUTAS.has(x)); if (f && r.rows[i][f]) api("abrir_ruta", r.rows[i][f]); },
      });
    }
    $("#ge-count").textContent = t("comun.resultados", { n: r.rows.length, m: r.limite });
  }

  function cambiosHechos() {
    buscar();
    document.dispatchEvent(new CustomEvent("agp:datos"));
  }

  /* ── Editor ──────────────────────────────────────────────────────────── */
  function editar(fila) {
    const r = datos, tb = tab, pk = fila[r.pk];
    const cuerpo = el(`<div class="form-grid"></div>`);
    r.editables.forEach(({ label, campo, opciones }) => {
      const val = fila[campo === "_pieza_malla" ? "pieza" : campo] || "";
      const full = /ruta/.test(campo) ? " full" : "";
      if (opciones) {
        const cur = opciones.includes(val) ? val : opciones[0];
        cuerpo.appendChild(el(`<div class="field${full}"><label>${esc(label)}</label><select class="input" data-c="${esc(campo)}">
          ${opciones.map((o) => `<option ${o === cur ? "selected" : ""}>${esc(o)}</option>`).join("")}</select></div>`));
      } else {
        cuerpo.appendChild(el(`<div class="field${full}"><label>${esc(label)}</label>
          <input class="input${full ? " input-ruta" : ""}" data-c="${esc(campo)}" value="${esc(val)}" spellcheck="false" /></div>`));
      }
    });
    if ("estado" in fila) {
      cuerpo.appendChild(el(`<div class="field full"><label>${esc(t("ge.estado"))}</label>
        <div><span class="estado ${esc(fila.estado || "NA")}">${esc(fila.estado || "—")}</span></div></div>`));
    }
    const pie = el(`<div style="display:contents"><span class="msg" data-msg></span><span class="sp"></span>
      ${r.anulable ? `<button class="btn btn-danger" data-a="anular">${icon("alert")}${esc(t("ge.anular"))}</button>` : ""}
      <button class="btn" data-a="cancel">${esc(t("comun.cancelar"))}</button>
      <button class="btn btn-primary" data-a="ok">${icon("save")}${esc(t("comun.guardar"))}</button></div>`);
    const m = abrirModal({ titulo: t("ge.editar"), sub: `${r.pk.toUpperCase()}: ${pk}`, icono: "edit", ancho: 680, cuerpo, pie });
    const msg = $("[data-msg]", pie);
    $$("input", cuerpo).forEach((inp) => inp.addEventListener("keydown", (e) => { if (e.key === "Enter") $('[data-a="ok"]', pie).click(); }));
    $('[data-a="cancel"]', pie).addEventListener("click", m.cerrar);
    $('[data-a="ok"]', pie).addEventListener("click", async (e) => {
      const valores = {};
      $$("[data-c]", cuerpo).forEach((n) => (valores[n.dataset.c] = n.value.trim()));
      e.currentTarget.disabled = true;
      const res = await api("gestion_guardar", tb, pk, valores);
      e.currentTarget.disabled = false;
      if (!res.ok) { msg.className = "msg err"; msg.textContent = "✘ " + res.error; return; }
      msg.className = "msg ok"; msg.textContent = "✔ " + t("ge.guardado");
      toast(`${pk} — ${t("ge.guardado")}`, "ok");
      setTimeout(m.cerrar, 900);
      cambiosHechos();
    });
    const bAn = $('[data-a="anular"]', pie);
    if (bAn) bAn.addEventListener("click", async () => {
      const ok = await confirmar(t("ge.anular_t"), t("ge.anular_txt", { pk: `${r.pk.toUpperCase()}: ${pk}` }),
        { ok: t("ge.anular_si"), peligro: true });
      if (!ok) return;
      const res = await api("gestion_anular", tb, pk);
      if (!res.ok) { msg.className = "msg err"; msg.textContent = "✘ " + res.error; return; }
      const partes = Object.entries(res.resultado || {}).filter(([, v]) => v > 0).map(([k, v]) => `${k}: ${v}`);
      msg.className = "msg warn"; msg.textContent = `✔ ${t("ge.anulado")} — ${partes.join(", ") || t("ge.sin_cambios")}`;
      toast(`${pk} — ${t("ge.anulado")}`, "warn");
      setTimeout(m.cerrar, 1300);
      cambiosHechos();
    });
  }

  /* ── Insertar manual ─────────────────────────────────────────────────── */
  function insertar() {
    const cuerpo = el(`<div><p class="dim" style="margin-bottom:12px">${esc(t("ins.elige"))}</p><div class="pick-list">
      <button class="pick" data-t="pasta_plata">${icon("coin")}<span><b>${esc(t("ins.pasta"))}</b><small>${esc(t("ins.pasta_d"))}</small></span></button>
      <button class="pick" data-t="vinilos">${icon("layers")}<span><b>${esc(t("ins.vinilos"))}</b><small>${esc(t("ins.vinilos_d"))}</small></span></button>
    </div></div>`);
    const m = abrirModal({ titulo: t("ins.titulo"), icono: "plus", color: "var(--ok-ink)", ancho: 460, cuerpo });
    $$(".pick", cuerpo).forEach((b) => b.addEventListener("click", () => { m.cerrar(); formulario(b.dataset.t); }));
  }

  async function formulario(tabla) {
    const r = await api("gestion_formulario", tabla);
    if (!r.ok) return alerta("error", t("ins.titulo"), r.error);
    const cuerpo = el(`<div class="form-grid"></div>`);
    r.campos.forEach(({ label, campo, auto, opciones }) => {
      if (opciones) {
        cuerpo.appendChild(el(`<div class="field"><label>${esc(label)}</label><select class="input" data-c="${esc(campo)}">
          ${opciones.map((o) => `<option>${esc(o)}</option>`).join("")}</select></div>`));
      } else {
        const full = /ruta/.test(campo) ? " full" : "";
        cuerpo.appendChild(el(`<div class="field${full}"><label>${esc(label)}${auto ? `<span class="badge-auto">${esc(t("ins.auto"))}</span>` : ""}</label>
          <input class="input" data-c="${esc(campo)}" ${auto ? `readonly value="${esc(r.consecutivo)}"` : ""} spellcheck="false" /></div>`));
      }
    });
    const pie = el(`<div style="display:contents"><span class="msg" data-msg></span><span class="sp"></span>
      <button class="btn" data-a="cancel">${esc(t("comun.cancelar"))}</button>
      <button class="btn btn-ok" data-a="ok">${icon("plus")}${esc(t("ins.insertar"))}</button></div>`);
    const titulo = tabla === "pasta_plata" ? t("ins.pasta") : t("ins.vinilos");
    const m = abrirModal({ titulo: `${t("ge.insertar")} — ${titulo}`, sub: t("ins.completa"), icono: tabla === "pasta_plata" ? "coin" : "layers",
      color: "var(--ok-ink)", ancho: 640, cuerpo, pie });
    const msg = $("[data-msg]", pie);
    $('[data-a="cancel"]', pie).addEventListener("click", m.cerrar);
    $('[data-a="ok"]', pie).addEventListener("click", async (e) => {
      const valores = {};
      $$("[data-c]:not([readonly])", cuerpo).forEach((n) => (valores[n.dataset.c] = n.value.trim()));
      e.currentTarget.disabled = true;
      const res = await api("gestion_insertar", tabla, valores);
      if (!res.ok) { e.currentTarget.disabled = false; msg.className = "msg err"; msg.textContent = "✘ " + String(res.error).slice(0, 120); return; }
      const ro = $("[readonly]", cuerpo); if (ro) ro.value = res.pk;
      msg.className = "msg ok"; msg.textContent = "✔ " + t("ins.insertado", { pk: res.pk });
      toast(t("ins.insertado", { pk: res.pk }), "ok");
      setTimeout(m.cerrar, 1500);
      cambiosHechos();
    });
  }

  /* ── Separar vitro / malla (3 pasos) ─────────────────────────────────── */
  function separar() {
    let rid = null, confirmado = false;
    const nv = numInput(1), ng = numInput(0), np = numInput(0);
    const campo = (k, lbl, full = "") => `<div class="field${full}"><label>${esc(lbl)}</label><input class="input" data-c="${k}" spellcheck="false" /></div>`;
    const cuerpo = el(`<div>
      <div class="sec-t">${esc(t("sep.p1"))}</div>
      <div class="qty-row"><div><b>${esc(t("asg.vitros"))}</b><small>T-xxxx</small></div><span data-n="v"></span></div>
      <div class="qty-row"><div><b>${esc(t("asg.grandes"))}</b><small>A-xxxx</small></div><span data-n="g"></span></div>
      <div class="qty-row" style="margin-bottom:16px"><div><b>${esc(t("asg.pequenas"))}</b><small>número</small></div><span data-n="p"></span></div>
      <div class="sec-t">${esc(t("sep.p2"))}</div>
      <div class="codes" style="margin-bottom:16px"><div class="dim" data-st>${esc(t("sep.vacio"))}</div><div data-codes></div></div>
      <div class="sec-t">${esc(t("sep.p3"))}</div>
      <div class="form-grid">
        ${campo("vehiculo", t("sep.vehiculo"))}${campo("cod_vehiculo", t("sep.cod"))}
        ${campo("version", t("sep.version"))}${campo("pieza", t("sep.pieza"))}
        ${campo("cod_completo", t("sep.cod_completo"))}
        <div class="field"><label>${esc(t("sep.bn"))}</label><select class="input" data-c="bnerig"><option>BN</option><option>BNI</option></select></div>
        <div class="field full"><label>${esc(t("sep.ruta"))}</label><div class="field-row">
          <input class="input input-ruta" data-c="ruta" spellcheck="false" /><button class="btn" type="button" data-a="pick">${icon("file")}</button></div></div>
        ${campo("responsable", t("sep.responsable"), " full")}
      </div></div>`);
    $('[data-n="v"]', cuerpo).replaceWith(nv); $('[data-n="g"]', cuerpo).replaceWith(ng); $('[data-n="p"]', cuerpo).replaceWith(np);
    const pie = el(`<div style="display:contents"><span class="sp"></span>
      <button class="btn btn-danger" data-a="cerrar">${icon("x")}<span>${esc(t("sep.cancelar"))}</span></button>
      <button class="btn btn-primary" data-a="res">${icon("lock")}<span>${esc(t("sep.reservar"))}</span></button>
      <button class="btn btn-ok" data-a="ok" disabled>${icon("check")}<span>${esc(t("sep.confirmar"))}</span></button></div>`);
    const m = abrirModal({
      titulo: t("sep.titulo"), sub: t("sep.sub"), icono: "lock", color: "#1f6feb", ancho: 700, cuerpo, pie, cerrarFuera: false,
      alCerrar: async () => { if (rid && !confirmado) await api("cancelar_reserva", rid); if (confirmado) cambiosHechos(); },
    });
    const st = $("[data-st]", cuerpo), codes = $("[data-codes]", cuerpo);
    const bRes = $('[data-a="res"]', pie), bOk = $('[data-a="ok"]', pie), bCerrar = $('[data-a="cerrar"]', pie);
    $('[data-a="pick"]', cuerpo).addEventListener("click", async () => {
      const r = await api("elegir_archivo", "ruta", "");
      if (r.ok && r.ruta) $('[data-c="ruta"]', cuerpo).value = r.ruta;
    });
    bCerrar.addEventListener("click", m.intentarCerrar);
    bRes.addEventListener("click", async () => {
      st.textContent = t("asg.reservando"); st.className = "dim"; codes.innerHTML = "";
      const r = await api("reservar", nv.valor(), ng.valor(), np.valor(), rid);
      if (!r.ok) { st.textContent = `Error: ${r.error}`; st.className = "msg err"; rid = null; bOk.disabled = true; return; }
      rid = r.rid;
      if (r.aviso) alerta("warning", t("sep.titulo"), r.aviso);
      codes.innerHTML = lineasCodigos(r); enlazarCopiar(codes);
      st.textContent = "✔ " + t("sep.reservado"); st.className = "msg ok";
      bOk.disabled = false;
      $("span", bRes).textContent = t("sep.rereservar");
    });
    bOk.addEventListener("click", async () => {
      const datosF = {};
      $$("[data-c]", cuerpo).forEach((n) => (datosF[n.dataset.c] = n.value.trim()));
      bOk.disabled = true;
      const r = await api("confirmar_separar", rid, datosF);
      if (!r.ok) { bOk.disabled = false; return alerta("error", t("sep.confirmar"), r.error); }
      confirmado = true;
      st.textContent = "✔ " + t("sep.asignado"); st.className = "msg ok";
      bRes.disabled = true; [nv, ng, np].forEach((n) => n.bloquear(true));
      bCerrar.className = "btn"; $("span", bCerrar).textContent = t("comun.cerrar");
      m.sub(`${t("ge.guardado")} · Veh: ${datosF.vehiculo || "—"} | Resp: ${datosF.responsable || "—"}`);
      toast(t("sep.asignado"), "ok");
    });
  }

  function refresh() { buscar(); }
  function init() {
    pintarTabs();
    $("#ge-q").addEventListener("input", debounce(buscar, 300));
    $("#ge-separar").addEventListener("click", separar);
    $("#ge-insertar").addEventListener("click", insertar);
  }
  return { init, refresh };
})();
