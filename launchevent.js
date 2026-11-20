/* Aviso al enviar (Smart Alerts, modo «PromptUser»): si el asunto no lleva código de obra, avisa y deja
   elegir entre cancelar el envío o «Enviar de todos modos» (para el ~15 % de correos que no son de obra). */

// Código al principio del asunto, tras los prefijos de respuesta/reenvío y etiquetas tipo [EXTERNO]:
// 0905-19 · 0905-29_PCI · 0157-123 · C9999 (confidencial)
var PREFIJOS = "(?:(?:\\[[^\\]]{1,30}\\]|(?:re|rv|res|fw|fwd|enc)\\s*:)\\s*)*";
var REGEX_CODIGO = new RegExp("^\\s*" + PREFIJOS + "(?:\\d{4}-\\d{2,3}(?:[_-][A-Za-z]+)?|C9999)(?=\\s|$|[-–—|:])", "i");
// Algo que parece un código pero está mal escrito: 599-111 · 0559.111 · 0559 111 · 0559_111
var REGEX_PARECIDO = new RegExp("^\\s*" + PREFIJOS + "\\d{1,4}\\s*[-._ ]\\s*\\d{1,3}(?=\\s|$|[-–—|:_])", "i");

var MENSAJE_FALTA = "Este correo no lleva código de obra en el asunto. " +
  "Pulsa «Elegir obra» para abrir la lista, o «Enviar de todos modos» si no corresponde a ninguna obra.";
var MENSAJE_PARECIDO = "El asunto empieza por algo que parece un código de obra, pero no está bien escrito (el formato es 0559-111). " +
  "Pulsa «Elegir obra» para corregirlo, o «Enviar de todos modos» si no es un código.";

// Botón propio del aviso: cancela el envío y abre el panel «Obra» (el Control btnObra del manifiesto).
function avisoFalta(parecido) {
  return { allowEvent: false, errorMessage: parecido ? MENSAJE_PARECIDO : MENSAJE_FALTA, cancelLabel: "Elegir obra", commandId: "btnObra" };
}

function onMessageSendHandler(event) {
  var terminado = false;
  function terminar(resultado) {
    if (terminado) return;
    terminado = true;
    event.completed(resultado);
  }
  // Si leer el asunto tarda o falla, no se deja a Outlook esperando: se avisa igualmente.
  var reloj = setTimeout(function () { terminar(avisoFalta()); }, 4000);
  try {
    Office.context.mailbox.item.subject.getAsync(function (resultado) {
      clearTimeout(reloj);
      if (resultado.status !== Office.AsyncResultStatus.Succeeded) {
        terminar({ allowEvent: true }); // si no se puede leer el asunto, no se molesta
        return;
      }
      var asunto = resultado.value || "";
      terminar(REGEX_CODIGO.test(asunto) ? { allowEvent: true } : avisoFalta(REGEX_PARECIDO.test(asunto)));
    });
  } catch (e) {
    clearTimeout(reloj);
    terminar(avisoFalta());
  }
}

Office.onReady(function () {});
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
