/* Pestaña Roles y Permisos (solo admin). */
"use strict";

const Roles = (() => {
  let usuarios = [], roles = [], sel = null, lista = [];

  async function cargar() {
    const cont = $("#ro-tabla");
    estadoCarga(cont);
    const r = await api("roles_listar");
    if (!r.ok) { estadoError(cont, r.error, cargar); return; }
    usuarios = r.usuarios; roles = r.roles;
    $("#ro-rol").innerHTML = `<option value="">${esc(t("ro.sin_rol"))}</option>` +
      roles.map((x) => `<option value="${esc(x)}">${esc(x)}</option>`).join("");
    pintar();
    if (sel) { const u = usuarios.find((x) => x.id === sel.id); if (u) elegir(u, false); }
  }

  function pintar() {
    const q = $("#ro-q").value.trim().toLowerCase();
    lista = usuarios.filter((u) => !q || u.nombre.toLowerCase().includes(q) || u.usuario.toLowerCase().includes(q));
    $("#ro-count").textContent = t("ro.usuarios", { n: lista.length });
    if (!lista.length) { estadoVacio($("#ro-tabla"), t("comun.sin_datos")); return; }
    pintarTabla($("#ro-tabla"), {
      headers: [t("ro.nombre"), t("ro.usuario"), t("ro.rol"), t("ro.estatus")],
      filas: lista.map((u) => [u.nombre, u.usuario, u.rol, u.estatus ? t("ro.activo") : t("ro.inactivo")]),
      claves: new Set(),
      render: (i, j) => {
        const u = lista[i];
        if (j === 2) return `<span class="rol-tag ${u.rol ? "" : "none"}">${esc(u.rol || t("ro.sin_rol"))}</span>`;
        if (j === 3) return `<span class="estado ${u.estatus ? "ASIGNADO" : "CANCELADO"}">${esc(u.estatus ? t("ro.activo") : t("ro.inactivo"))}</span>`;
        return null;
      },
      alDobleClic: (_f, i) => elegir(lista[i]),
      alEnter: (_f, i) => elegir(lista[i]),
    });
  }

  function elegir(u, limpiarMsg = true) {
    sel = u;
    $("#ro-nombre").textContent = u.nombre || "—";
    $("#ro-usuario").textContent = u.usuario || "—";
    $("#ro-rol").value = u.rol || "";
    $("#ro-est").checked = !!u.estatus;
    $("#ro-save").disabled = false;
    if (limpiarMsg) $("#ro-msg").textContent = "";
  }

  async function guardar() {
    const msg = $("#ro-msg");
    if (!sel) { msg.className = "msg err"; msg.textContent = t("ro.sel"); return; }
    const r = await api("roles_guardar", sel.id, $("#ro-rol").value, $("#ro-est").checked, sel.usuario);
    if (!r.ok) { msg.className = "msg err"; msg.textContent = "Error: " + r.error; return; }
    msg.className = "msg ok"; msg.textContent = "✔ " + t("ro.guardado");
    toast(`${sel.nombre} — ${t("ro.guardado")}`, "ok");
    cargar();
  }

  function init() {
    $("#ro-q").addEventListener("input", debounce(pintar, 150));
    $("#ro-reload").addEventListener("click", cargar);
    $("#ro-save").addEventListener("click", guardar);
    $("#ro-tabla").addEventListener("click", (e) => {
      const tr = e.target.closest("tr[data-i]");
      if (tr && lista[+tr.dataset.i]) elegir(lista[+tr.dataset.i]);
    });
  }
  return { init, refresh: cargar };
})();
