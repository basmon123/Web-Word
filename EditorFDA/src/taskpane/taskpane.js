/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    
    // Asignar eventos a los botones (Aquí conectamos el HTML con el JS)
    document.getElementById("btnGenerar").onclick = generarDocumento;
    document.getElementById("btnLimpiar").onclick = limpiarFormato;
    document.getElementById("btnFecha").onclick = insertarFecha;
    
    document.getElementById("btnTitulo1").onclick = () => aplicarEstilo("Heading 1");
    document.getElementById("btnTitulo2").onclick = () => aplicarEstilo("Heading 2");
    document.getElementById("btnTitulo3").onclick = () => aplicarEstilo("Heading 3");

    console.log("Taskpane cargado y botones listos.");
  }
});

// --- FUNCIÓN 1: GENERAR DOCUMENTO (Rellena los Content Controls) ---
async function generarDocumento() {
  return Word.run(async (context) => {
    // 1. Obtener valores de los inputs del HTML
    const datos = {
      "Cliente": document.getElementById("inCliente").value,
      "Division": document.getElementById("inDivision").value,
      "Proyecto": document.getElementById("inProyecto").value,
      "Contrato": document.getElementById("inContrato").value,
      "API": document.getElementById("inAPI").value,
      "Servicios": document.getElementById("inServicios").value,
      "NombreDoc": document.getElementById("inNombreDoc").value,
      "Codigo": document.getElementById("inCodigo").value,
      "Revision": document.getElementById("inRevision").value
    };

    // 2. Buscar Content Controls por su "Title" o "Tag" y reemplazar texto
    // Nota: Esto asume que en tu Word tienes controles con estos títulos.
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();

    for (let i = 0; i < contentControls.items.length; i++) {
      const cc = contentControls.items[i];
      // Si el título del control coincide con alguna llave de nuestros datos
      if (datos[cc.title]) {
        cc.insertText(datos[cc.title], "Replace");
      }
    }

    document.getElementById("mensajeEstado").innerText = "Documento actualizado con éxito.";
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

// --- FUNCIÓN 2: LIMPIAR FORMATO ---
async function limpiarFormato() {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.clear(); // Borra formato directo (negrita, colores manuales, etc.)
    await context.sync();
  });
}

// --- FUNCIÓN 3: INSERTAR FECHA ---
async function insertarFecha() {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    const hoy = new Date().toLocaleDateString("es-ES", {
        year: 'numeric', month: 'long', day: 'numeric'
    });
    range.insertText(hoy, "Replace");
    await context.sync();
  });
}

// --- FUNCIÓN 4: APLICAR ESTILOS (Títulos) ---
async function aplicarEstilo(nombreEstilo) {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.style = nombreEstilo;
    await context.sync();
  });
}