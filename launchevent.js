/* Aviso al enviar (Smart Alerts, modo «PromptUser»): si el asunto no lleva código de obra, avisa y deja
   elegir entre cancelar el envío o «Enviar de todos modos» (para el ~15 % de correos que no son de obra). */

// Interruptor: con true, además de comprobar el formato se comprueba que el código exista en la lista de obras.
// Se activará cuando la lista sea la real (SharePoint); con la lista de demostración avisaría de todos los códigos reales.
var COMPROBAR_EXISTENCIA = false;
var URL_OBRAS = "obras_demo.json";   // misma fuente que el panel (taskpane.js)

// Código al principio del asunto, tras los prefijos de respuesta/reenvío y etiquetas tipo [EXTERNO]:
// 0905-19 · 0905-29_PCI · 0157-123 · C9999 (confidencial)
var PREFIJOS = "(?:(?:\\[[^\\]]{1,30}\\]|(?:re|rv|res|fw|fwd|enc|tr)\\s*:)\\s*)*";
var REGEX_CODIGO = new RegExp("^\\s*" + PREFIJOS + "(\\d{4}-\\d{2,3}(?:[_-][A-Za-z]+)?|C9999)(?=\\s|$|[-–—|:])", "i");
// Algo que parece un código pero está mal escrito: 599-111 · 0559.111 · 0559 111 · 0559_111 · 0559107 (pegado)
var REGEX_PARECIDO = new RegExp("^\\s*" + PREFIJOS + "(?:\\d{1,4}\\s*[-._ ]\\s*\\d{1,3}|\\d{7})(?=\\s|$|[-–—|:_])", "i");

var MENSAJE_FALTA = "Este correo no lleva código de obra en el asunto. " +
  "Pulsa «Elegir obra» para abrir la lista, o «Enviar de todos modos» si no corresponde a ninguna obra.";
var MENSAJE_PARECIDO = "El asunto empieza por algo que parece un código de obra, pero no está bien escrito (el formato es 0559-111). " +
  "Pulsa «Elegir obra» para corregirlo, o «Enviar de todos modos» si no es un código.";

function mensajeInexistente(codigo) {
  return "El código " + codigo + " no está en la lista de obras. " +
    "Pulsa «Elegir obra» para comprobarlo, o «Enviar de todos modos» si es correcto.";
}

// Botón propio del aviso: cancela el envío y abre el panel «Obra» (el Control btnObra del manifiesto).
function aviso(mensaje) {
  return { allowEvent: false, errorMessage: mensaje, cancelLabel: "Elegir obra", commandId: "btnObra" };
}

function onMessageSendHandler(event) {
  var terminado = false;
  function terminar(resultado) {
    if (terminado) return;
    terminado = true;
    event.completed(resultado);
  }
  // Si leer el asunto tarda o falla, no se deja a Outlook esperando: se avisa igualmente.
  var reloj = setTimeout(function () { terminar(aviso(MENSAJE_FALTA)); }, 4000);
  try {
    Office.context.mailbox.item.subject.getAsync(function (resultado) {
      clearTimeout(reloj);
      if (resultado.status !== Office.AsyncResultStatus.Succeeded) {
        terminar({ allowEvent: true }); // si no se puede leer el asunto, no se molesta
        return;
      }
      var asunto = resultado.value || "";
      var m = asunto.match(REGEX_CODIGO);
      if (!m) { terminar(aviso(REGEX_PARECIDO.test(asunto) ? MENSAJE_PARECIDO : MENSAJE_FALTA)); return; }
      var codigo = m[1].toUpperCase();
      if (!COMPROBAR_EXISTENCIA || codigo === "C9999") { terminar({ allowEvent: true }); return; }
      // Comprobar que el código existe. Si la lista no se puede leer, no se molesta: se deja enviar.
      setTimeout(function () { terminar({ allowEvent: true }); }, 3000);
      fetch(URL_OBRAS, { cache: "no-store" }).then(function (r) { return r.json(); }).then(function (obras) {
        var existe = obras.some(function (o) { return String(o.codigo).toUpperCase() === codigo; });
        terminar(existe ? { allowEvent: true } : aviso(mensajeInexistente(codigo)));
      }).catch(function () { terminar({ allowEvent: true }); });
    });
  } catch (e) {
    clearTimeout(reloj);
    terminar(aviso(MENSAJE_FALTA));
  }
}

Office.onReady(function () {});
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
