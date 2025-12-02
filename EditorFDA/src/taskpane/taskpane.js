/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const btn = document.getElementById("btnGenerar");
    if (btn) {
      btn.onclick = run;
    }
  }
});

// --- FUNCIÓN DEL PANEL (Generador) ---
async function run() {
  try {
    const getVal = (id) => document.getElementById(id) ? document.getElementById(id).value : "";
    
    // Captura de datos
    const vCliente   = getVal("inCliente");
    const vDivision  = getVal("inDivision");
    const vProyecto  = getVal("inProyecto");
    const vContrato  = getVal("inContrato");
    const vAPI       = getVal("inAPI");
    const vServicios = getVal("inServicios");
    const vNombreDoc = getVal("inNombreDoc");
    const vCodigo    = getVal("inCodigo");
    const vRevision  = getVal("inRevision");

    const msgLabel = document.getElementById("mensajeEstado");
    if (msgLabel) msgLabel.textContent = "Procesando...";

    await Word.run(async (context) => {
      const mapaDeTags = [
        { tag: "ccCliente",       valor: vCliente },
        { tag: "ccDivisión",      valor: vDivision },
        { tag: "ccServicios",     valor: vServicios },
        { tag: "ccContrato",      valor: vContrato },
        { tag: "ccAPI",           valor: vAPI },
        { tag: "ccProyecto",      valor: vProyecto },
        { tag: "ccNombreDoc",     valor: vNombreDoc }, // Ojo: Verifica si tu tag es ccNombreDoc o ccNombre doc en Word
        { tag: "ccCliente_encabezado",   valor: vCliente },
        { tag: "ccD_encabezado",         valor: vDivision },
        { tag: "ccNProyecto_Encabezado", valor: vProyecto },
        { tag: "ccCodigo",               valor: vCodigo },
        { tag: "ccRevision",             valor: vRevision }
      ];

      let contadores = 0;
      for (let item of mapaDeTags) {
        let ccs = context.document.contentControls.getByTag(item.tag);
        ccs.load("items");
        await context.sync();

        if (ccs.items.length > 0) {
           for (let cc of ccs.items) {
             cc.insertText(item.valor, "Replace");
             contadores++;
           }
        }
      }

      await context.sync();
      if (msgLabel) msgLabel.textContent = "¡Listo! " + contadores + " campos actualizados.";
    });
  } catch (error) {
    console.error(error);
  }
}

// --- NUEVAS FUNCIONES DE LOS BOTONES (Comandos) ---

// Botón 1: Limpieza FDA (VERSIÓN SECUENCIAL SEGURA)
async function limpiarFormato(event) {
  try {
    await Word.run(async (context) => {
      // 1. Obtener selección
      const selection = context.document.getSelection();

      // --- FASE 1: ARREGLAR FUENTE ---
      // Cargamos solo la fuente primero
      context.load(selection, "font");
      await context.sync();

      selection.font.name = "Arial";
      selection.font.size = 11;
      selection.font.color = "black";
      selection.font.bold = false; // Quitamos negrita por si acaso

      // Guardamos cambios de fuente
      await context.sync();

      // --- FASE 2: ARREGLAR PÁRRAFO ---
      // Ahora cargamos el párrafo
      context.load(selection, "paragraphFormat");
      await context.sync();

      selection.paragraphFormat.alignment = "Justified";

      // Guardamos cambios de párrafo
      await context.sync();
    });
  } catch (error) {
    console.error("Error al formatear:", error);
  }
  
  if (event) event.completed();
}

// Botón 2: Insertar Fecha
async function insertarFecha(event) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const fechaHoy = new Date().toLocaleDateString();
    selection.insertText(fechaHoy, "Replace");
    await context.sync();
  });
  
  if (event) event.completed();
}

// --- REGISTRO ---
Office.actions.associate("limpiarFormato", limpiarFormato);
Office.actions.associate("insertarFecha", insertarFecha);