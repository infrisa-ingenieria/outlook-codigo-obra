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

// Solo se aceptan enlaces https (los datos vienen de fuera: nada de javascript: ni similares).
function enlaceSeguro(u) { return /^https:\/\//i.test(u || "") ? u : ""; }

function enlacesHtml(o) {
  var sp = enlaceSeguro(o.url_sharepoint), fo = enlaceSeguro(o.url_forma);
  var h = "";
  if (sp) h += '<a href="' + esc(sp) + '" target="_blank" rel="noopener" title="Abrir la carpeta de la obra en SharePoint">SP</a>';
  if (fo) h += '<a href="' + esc(fo) + '" target="_blank" rel="noopener" title="Abrir la obra en Autodesk Forma">Forma</a>';
  return h ? '<span class="enlaces">' + h + "</span>" : "";
}

function filaHtml(o) {
  var fav = favoritas.indexOf(o.codigo) >= 0;
  return '<div class="fila" data-codigo="' + esc(o.codigo) + '">' +
    '<button class="estrella" title="Favorita" data-fav="' + esc(o.codigo) + '">' + (fav ? "★" : "☆") + "</button>" +
    '<span class="codigo">' + codigoHtml(o.codigo) + "</span>" +
    '<span class="nombre" title="' + esc(o.nombre) + '">' + esc(o.nombre) + "</span>" +
    enlacesHtml(o) + "</div>";
}

// Código que ya lleva el asunto (tras RE:/FW:), p. ej. «0905-29_PCI»; "" si no lleva.
function codigoDelAsunto(asunto) {
  var m = (asunto || "").match(/^\s*(?:(?:re|rv|res|fw|fwd|enc)\s*:\s*)*(\d{4}-\d{2,3}(?:[_-][A-Za-z]+)?)(?=\s|$|[-–—|:])/i);
  return m ? m[1] : "";
}

// Franja superior: qué código lleva ahora el correo (útil al responder/reenviar) y enlaces a su carpeta.
function mostrarActual(asunto) {
  var caja = document.getElementById("actual");
  var cod = codigoDelAsunto(asunto);
  if (!cod) { caja.innerHTML = ""; return; }
  var o = porCodigo(cod);
  if (!o) {
    caja.innerHTML = "Este correo ya lleva <b>" + esc(cod) + '</b><span class="nota">Ese código no está en la lista de obras: compruébalo.</span>';
    return;
  }
  caja.innerHTML = "Este correo ya es de <b>" + esc(o.codigo) + "</b> · " + esc(o.nombre) + " " + enlacesHtml(o) +
    '<span class="nota">Elige otra obra abajo solo si quieres cambiarlo.</span>';
}

function refrescarActual() {
  try {
    Office.context.mailbox.item.subject.getAsync(function (r) {
      mostrarActual(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "");
    });
  } catch (e) { mostrarActual(""); }
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
        mostrarActual(nuevo);
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
    if (e.target.closest("a")) return;   // enlaces SP / Forma: se abren, no cambian el asunto
    var fav = e.target.closest("[data-fav]");
    if (fav) { e.stopPropagation(); alternarFavorita(fav.getAttribute("data-fav")); return; }
    var fila = e.target.closest(".fila");
    if (fila) ponerEnAsunto(fila.getAttribute("data-codigo"));
  });
  document.getElementById("buscar").addEventListener("input", pintar);

  cargarObras().then(function () { pintar(); refrescarActual(); }).catch(function () { estado("No se pudo cargar la lista de obras."); });
});
