/* Arranque: login / SSO, navegación, banner de BD, tracking de actividad. */
"use strict";

const NAV_ICON = { arte: "palette", bd: "database", gestion: "edit", scanner: "scan", roles: "users", logos: "laser" };
const MODULOS = { arte: Arte, bd: BD, gestion: Gestion, scanner: Scanner, roles: Roles, logos: Logos };
let USUARIO = null, PAGINA = null;
const pendientesRefresco = new Set();

/* ── Login ─────────────────────────────────────────────────────────────── */
function mostrarLogin() {
  $("#login").classList.remove("hide");
  setTimeout(() => $("#lg-user").focus(), 50);
  $("#login-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    const msg = $("#lg-msg"), btn = $("#lg-btn");
    const u = $("#lg-user").value.trim(), p = $("#lg-pass").value.trim();
    if (!u || !p) { msg.className = "login-msg"; msg.textContent = t("login.campos"); return; }
    btn.disabled = true; btn.textContent = t("login.verificando"); msg.textContent = "";
    const r = await api("login", u, p);
    btn.disabled = false; btn.textContent = t("login.entrar");
    if (!r.ok) {
      msg.className = "login-msg";
      msg.textContent = r.error === "credenciales" ? t("login.malas")
        : r.error === "campos" ? t("login.campos") : `${t("login.conexion")}: ${r.detalle || r.error}`;
      $("#lg-pass").select();
      return;
    }
    msg.className = "login-msg ok";
    msg.textContent = t("login.bienvenido", { n: (r.usuario.nombre || "").split(" ")[0] });
    setTimeout(() => { $("#login").classList.add("hide"); entrar(r.usuario); }, 350);
  });
}

/* ── Entrar a la app ───────────────────────────────────────────────────── */
function entrar(u) {
  USUARIO = u;
  if (!u.rol) {
    $("#sin-rol-txt").textContent = t("sinrol.txt", { n: (u.nombre || "").split(" ")[0] });
    $("#sin-rol").classList.remove("hide");
    return;
  }
  const nombre = u.nombre || u.usuario || "";
  $("#u-nombre").textContent = nombre.split(" ").slice(0, 2).join(" ");
  $("#u-rol").textContent = u.rol;
  $("#u-avatar").textContent = nombre.split(" ").filter(Boolean).slice(0, 2).map((x) => x[0]).join("").toUpperCase() || "?";
  pintarNav();
  Object.values(MODULOS).forEach((m) => m.init());
  $("#app").classList.remove("hide");
  pintarIconos();
  aplicarI18n();
  setTema(temaActual());
  if (u.paginas.length) mostrar(u.paginas[0].id);
  probarConexion();
  pollLoop();
  // Clic / tecla real = actividad (el módulo de tracking ya limita a 1 cada 30 s)
  let ultima = 0;
  const act = () => { const ahora = Date.now(); if (ahora - ultima > 25000) { ultima = ahora; api("actividad"); } };
  document.addEventListener("mousedown", act, true);
  document.addEventListener("keydown", act, true);
}

function pintarNav() {
  $("#sb-nav").innerHTML = USUARIO.paginas.map((p, i) =>
    `<button class="sb-item" data-p="${p.id}" title="${esc(t("pag." + p.id))}">${icon(NAV_ICON[p.id])}
      <span>${esc(t("pag." + p.id))}</span><span class="sb-key">Alt+${i + 1}</span></button>`).join("");
  $$("#sb-nav .sb-item").forEach((b) => b.addEventListener("click", () => mostrar(b.dataset.p)));
  if (PAGINA) $$("#sb-nav .sb-item").forEach((b) => b.classList.toggle("active", b.dataset.p === PAGINA));
}

function mostrar(id) {
  if (!USUARIO.paginas.some((p) => p.id === id)) return;
  const primeraVez = !$(`.page[data-page="${id}"]`).dataset.visto;
  PAGINA = id;
  $$(".page").forEach((s) => s.classList.toggle("on", s.dataset.page === id));
  $$("#sb-nav .sb-item").forEach((b) => b.classList.toggle("active", b.dataset.p === id));
  $("#tb-title").textContent = t("pag." + id);
  $("#tb-sub").textContent = t("sub." + id);
  $("#content").scrollTop = 0;
  const nombreTracking = USUARIO.paginas.find((p) => p.id === id).nombre;
  api("cambio_pestana", nombreTracking);
  $(`.page[data-page="${id}"]`).dataset.visto = "1";
  // Igual que antes: Consultar / Gestión se refrescan siempre al entrar
  if (id === "bd" || id === "gestion" || pendientesRefresco.has(id) || (primeraVez && id === "roles")) {
    pendientesRefresco.delete(id);
    MODULOS[id].refresh();
  }
  if (MODULOS[id].alMostrar) MODULOS[id].alMostrar();
}

/* ── Sidebar: colapsado, se abre al pasar el mouse, botón para fijarlo ── */
function initSidebar() {
  const app = $("#app"), sb = $("#sidebar");
  let fijo = !!Prefs.get("menu_fijo", false);
  const aplicar = () => {
    app.classList.toggle("pinned", fijo);
    $("#sb-pin-ic").innerHTML = icon(fijo ? "pinoff" : "pin");
    $('#sb-pin [data-i18n]').dataset.i18n = fijo ? "nav.soltar" : "nav.fijar";
    $('#sb-pin [data-i18n]').textContent = t(fijo ? "nav.soltar" : "nav.fijar");
  };
  sb.addEventListener("mouseenter", () => sb.classList.add("open"));
  sb.addEventListener("mouseleave", () => sb.classList.remove("open"));
  $("#sb-pin").addEventListener("click", () => {
    fijo = !fijo;
    Prefs.set("menu_fijo", fijo);
    aplicar();
  });
  aplicar();
}

/* ── Conexión a BD: pastilla + banner con Reintentar ───────────────────── */
async function probarConexion() {
  const pill = $("#pill-bd");
  pill.className = "pill"; $("#pill-bd-txt").textContent = t("comun.cargando");
  const r = await api("test_conexion");
  if (r.ok) {
    pill.className = "pill ok"; $("#pill-bd-txt").textContent = t("bd.ok");
    $("#banner").classList.add("hide");
  } else {
    pill.className = "pill err"; $("#pill-bd-txt").textContent = t("bd.err");
    $("#banner-txt").textContent = t("bd.sin_conexion");
    $("#banner-det").textContent = `${t("bd.sin_conexion_d")}  ·  ${String(r.error).slice(0, 160)}`;
    $("#banner").classList.remove("hide");
  }
}

/* ── Arranque ──────────────────────────────────────────────────────────── */
async function arrancar() {
  const r = await api("inicio");
  Prefs.cargar(r.prefs);
  LANG = Prefs.get("lang", "es");
  setTema(Prefs.get("tema", "dark"), false);
  pintarIconos();
  $("#banner-ic").innerHTML = icon("alert");
  $("#banner-x").innerHTML = icon("x");
  aplicarI18n();
  $$(".js-tema").forEach((b) => b.addEventListener("click", () => setTema(temaActual() === "dark" ? "light" : "dark")));
  $$("[data-lang]").forEach((b) => b.addEventListener("click", () => setLang(b.dataset.lang)));
  $("#banner-retry").addEventListener("click", probarConexion);
  $("#banner-x").addEventListener("click", () => $("#banner").classList.add("hide"));
  initSidebar();

  document.addEventListener("agp:lang", () => {
    if (!USUARIO) return;
    pintarNav();
    if (PAGINA) { $("#tb-title").textContent = t("pag." + PAGINA); $("#tb-sub").textContent = t("sub." + PAGINA); }
    setTema(temaActual());
  });
  // Cambios en BD (Gestión / arte creado) → refrescar las otras pestañas
  const marcarSucio = (ids) => ids.forEach((id) => { if (id === PAGINA) MODULOS[id].refresh(); else pendientesRefresco.add(id); });
  document.addEventListener("agp:datos", () => marcarSucio(["bd"]));
  Bus.on("arte_done", () => marcarSucio(["bd", "gestion"]));

  document.addEventListener("keydown", (e) => {
    if (!USUARIO || !e.altKey) return;
    const n = parseInt(e.key, 10);
    if (n >= 1 && n <= USUARIO.paginas.length) { e.preventDefault(); mostrar(USUARIO.paginas[n - 1].id); }
  });

  $("#sb-ver").textContent = `v${r.version || "3.0"} · Ingeniería`;
  if (r.ok && r.usuario) entrar(r.usuario);
  else mostrarLogin();
}

let _arrancado = false;
const _arrancarUnaVez = () => { if (!_arrancado) { _arrancado = true; arrancar(); } };
window.addEventListener("pywebviewready", _arrancarUnaVez);
if (window.pywebview && window.pywebview.api && window.pywebview.api.inicio) _arrancarUnaVez();
