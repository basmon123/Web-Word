/* global Office, Word */

Office.onReady(() => {
  // Inicialización si fuera necesaria
});

let dialog; // Variable para guardar la ventana

// 1. ESTA FUNCIÓN LA LLAMA EL BOTÓN DE LA CINTA
function abrirCatalogo(event) {
  // URL de tu archivo catalog.html (Asegúrate que coincida con tu GitHub/Localhost)
  // TRUCO: Usamos location.origin para que funcione en Local y en Nube sin cambiar código
  const url = window.location.origin + "/src/catalog/catalog.html";

  // Abrimos la ventana emergente
  Office.context.ui.displayDialogAsync(url, { height: 60, width: 50 }, 
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error.message);
      } else {
        dialog = asyncResult.value;
        // Nos ponemos a escuchar mensajes de la ventana
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, procesarMensaje);
      }
    }
  );
  
  if(event) event.completed();
}

// 2. ESTA FUNCIÓN RECIBE LA ORDEN DE LA VENTANA
async function procesarMensaje(arg) {
  dialog.close(); // Cerramos la ventana primero
  
  const mensaje = JSON.parse(arg.message); // Leemos el JSON {accion, plantilla, datos}

  if (mensaje.accion === "CREAR_DOCUMENTO") {
      await crearDocumentoNuevo(mensaje.plantilla, mensaje.datos);
  }
}

// 3. ESTA FUNCIÓN CREA EL WORD (Simulado por ahora)
async function crearDocumentoNuevo(nombrePlantilla, datosProyecto) {
  await Word.run(async (context) => {
    // AQUI OCURRE LA MAGIA: Creamos un documento nuevo en blanco
    // (A futuro, aquí cargaremos el Base64 de la plantilla real)
    const newDoc = context.application.createDocument();
    
    // Escribimos en el documento nuevo
    const body = newDoc.body;
    body.insertParagraph("DOCUMENTO GENERADO: " + nombrePlantilla.toUpperCase(), "Start");
    body.insertParagraph("Proyecto: " + datosProyecto.id + " - " + datosProyecto.nombre, "End");
    body.insertParagraph("Cliente: " + datosProyecto.cliente, "End");
    
    // Abrimos el documento nuevo
    newDoc.open();
    
    await context.sync();
  });
}

// REGISTRO OBLIGATORIO (Para que el XML encuentre la función)
// Nota: A veces se necesita 'g' o 'window' dependiendo del entorno, esto suele funcionar:
const g = typeof globalThis !== "undefined" ? globalThis : window;
g.abrirCatalogo = abrirCatalogo;