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

// Botón 1: Limpieza FDA (Versión FINAL: Fuente + Justificado Seguro)
async function limpiarFormato(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // FASE 1: FUENTE (Lo que ya funciona)
      context.load(selection, "font");
      await context.sync();

      selection.font.set({
        name: "Arial",
        size: 11,
        color: "#000000",
        bold: false,
        italic: false
      });
      
      await context.sync(); // Guardamos fuente

      // FASE 2: ALINEACIÓN (Blindada)
      // Cargamos específicamente el formato de párrafo ahora
      context.load(selection, "paragraphFormat");
      await context.sync();

      try {
        // Intentamos justificar
        selection.paragraphFormat.alignment = "Justified";
        await context.sync(); // Guardamos alineación
      } catch (alignError) {
        // Si falla la alineación (ej: dentro de una tabla compleja), 
        // solo lo ignoramos y seguimos, pero la fuente ya quedó arreglada.
        console.warn("No se pudo justificar, pero la fuente se cambió.");
      }
    });
  } catch (error) {
    console.error("Error General FDA:", error);
  } finally {
    // VITAL: Esto apaga el mensaje de "Trabajando..."
    if (event) event.completed();
  }
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