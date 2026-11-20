/* Panel «Obra»: favoritas, últimas y búsqueda; al elegir una obra pone «código – texto» en el asunto.
   Favoritas y últimas se guardan en roamingSettings (se sincronizan entre los Outlook de la misma cuenta). */

// De dónde salen las obras. v0: fichero de muestra. Después: URL del flujo de Power Automate que lee SharePoint.
var URL_OBRAS = "obras_demo.json";
var MAX_RECIENTES = 8;
var MAX_RESULTADOS = 40;

var obras = [];            // [{codigo, nombre}]
var favoritas = [];        // [codigo]
var recientes = [];        // [codigo]
var ajustes = null;

function sinAcentos(t) { return (t || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase(); }
function esc(t) { return String(t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

// Códigos con anexo: «0905-29_PCI» se muestra como 0905-29 con «PCI» debajo, igual que la app.
function codigoHtml(c) {
  var i = c.indexOf("_");
  return i < 0 ? esc(c) : esc(c.slice(0, i)) + "<small>" + esc(c.slice(i + 1)) + "</small>";
}

function guardar() {
  ajustes.set("favoritas", JSON.stringify(favoritas));
  ajustes.set("recientes", JSON.stringify(recientes));
  ajustes.saveAsync(function () {});
}

function leerLista(clave) {
  try { return JSON.parse(ajustes.get(clave) || "[]"); } catch (e) { return []; }
}

function porCodigo(c) { return obras.find(function (o) { return o.codigo === c; }); }

function filaHtml(o) {
  var fav = favoritas.indexOf(o.codigo) >= 0;
  return '<div class="fila" data-codigo="' + esc(o.codigo) + '">' +
    '<button class="estrella" title="Favorita" data-fav="' + esc(o.codigo) + '">' + (fav ? "★" : "☆") + "</button>" +
    '<span class="codigo">' + codigoHtml(o.codigo) + "</span>" +
    '<span class="nombre" title="' + esc(o.nombre) + '">' + esc(o.nombre) + "</span></div>";
}

function seccion(titulo, lista) {
  return lista.length ? "<h2>" + titulo + "</h2>" + lista.map(filaHtml).join("") : "";
}

function pintar() {
  var q = sinAcentos(document.getElementById("buscar").value).trim();
  var cont = document.getElementById("listas");
  if (!q) {
    var f = favoritas.map(porCodigo).filter(Boolean);
    var r = recientes.map(porCodigo).filter(Boolean).filter(function (o) { return favoritas.indexOf(o.codigo) < 0; });
    cont.innerHTML = seccion("Favoritas", f) + seccion("Últimas usadas", r) ||
      '<p class="vacio">Escribe un código, un cliente o un lugar para buscar.</p>';
    if (!f.length && !r.length) cont.innerHTML = '<p class="vacio">Escribe un código, un cliente o un lugar para buscar. Las que uses aparecerán aquí.</p>';
    return;
  }
  var palabras = q.split(/\s+/);
  var res = obras.filter(function (o) {
    var t = sinAcentos(o.codigo + " " + o.nombre);
    return palabras.every(function (p) { return t.indexOf(p) >= 0; });
  });
  var corto = res.slice(0, MAX_RESULTADOS);
  cont.innerHTML = corto.length ? seccion("Resultados (" + res.length + ")", corto) : '<p class="vacio">Ninguna obra coincide.</p>';
}

function estado(t) { document.getElementById("estado").textContent = t || ""; }

// Pone «código – texto» respetando prefijos RE:/RV:/FW: y sustituyendo un código anterior si lo hubiera.
function ponerEnAsunto(codigo) {
  var item = Office.context.mailbox.item;
  item.subject.getAsync(function (r) {
    var actual = r.status === Office.AsyncResultStatus.Succeeded ? (r.value || "") : "";
    var m = actual.match(/^\s*((?:(?:re|rv|res|fw|fwd|enc)\s*:\s*)*)(.*)$/i);
    var prefijo = m ? m[1] : "";
    var resto = m ? m[2] : actual;
    resto = resto.replace(/^\d{4}-\d{2,3}(?:[_-][A-Za-z]+)?\s*[-–—|:]?\s*/, "");   // quita el código anterior
    var nuevo = prefijo + codigo + (resto ? " – " + resto : " – ");
    item.subject.setAsync(nuevo, function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        recientes = [codigo].concat(recientes.filter(function (c) { return c !== codigo; })).slice(0, MAX_RECIENTES);
        guardar();
        estado("Asunto: " + nuevo);
        pintar();
      } else {
        estado("No se pudo cambiar el asunto.");
      }
    });
  });
}

function alternarFavorita(codigo) {
  var i = favoritas.indexOf(codigo);
  if (i >= 0) favoritas.splice(i, 1); else favoritas.push(codigo);
  guardar();
  pintar();
}

function cargarObras() {
  return fetch(URL_OBRAS, { cache: "no-store" }).then(function (r) { return r.json(); }).then(function (datos) {
    obras = datos.sort(function (a, b) { return a.codigo < b.codigo ? -1 : 1; });
  });
}

Office.onReady(function () {
  ajustes = Office.context.roamingSettings;
  favoritas = leerLista("favoritas");
  recientes = leerLista("recientes");

  document.getElementById("listas").addEventListener("click", function (e) {
    var fav = e.target.closest("[data-fav]");
    if (fav) { e.stopPropagation(); alternarFavorita(fav.getAttribute("data-fav")); return; }
    var fila = e.target.closest(".fila");
    if (fila) ponerEnAsunto(fila.getAttribute("data-codigo"));
  });
  document.getElementById("buscar").addEventListener("input", pintar);

  cargarObras().then(pintar).catch(function () { estado("No se pudo cargar la lista de obras."); });
});
