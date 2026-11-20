/* PRUEBA v2: responde al instante, sin leer el asunto, para comprobar que el aviso al enviar arranca. */

function onMessageSendHandler(event) {
  event.completed({
    allowEvent: false,
    errorMessage: "PRUEBA v2: el complemento responde. Aqui ira el aviso de falta de codigo de obra."
  });
}

Office.onReady(function () {});
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
