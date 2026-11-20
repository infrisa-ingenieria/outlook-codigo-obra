/* Aviso al enviar (Smart Alerts, modo «PromptUser»): si el asunto no lleva código de obra, avisa y deja
   elegir entre cancelar el envío o «Enviar de todos modos» (para el ~15 % de correos que no son de obra). */

// Código al principio del asunto, tras los prefijos de respuesta/reenvío: 0905-19 · 0905-29_PCI · 0157-123
var REGEX_CODIGO = /^\s*((re|rv|res|fw|fwd|enc)\s*:\s*)*\d{4}-\d{2,3}([_-][A-Za-z]+)?(\s|$|[-–—|:])/i;

function onMessageSendHandler(event) {
  Office.context.mailbox.item.subject.getAsync(function (resultado) {
    if (resultado.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true }); // si no se puede leer el asunto, no se molesta
      return;
    }
    var asunto = resultado.value || "";
    if (REGEX_CODIGO.test(asunto)) {
      event.completed({ allowEvent: true });
    } else {
      event.completed({
        allowEvent: false,
        errorMessage: "Este correo no lleva código de obra en el asunto (formato: 0000-00 – texto). " +
                      "Pulsa «Obra» en la cinta para elegirlo, o «Enviar de todos modos» si no corresponde a ninguna obra."
      });
    }
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
