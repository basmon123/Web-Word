/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const btn = document.getElementById("btnGenerar");
    if (btn) {
      btn.onclick = run;
    }
  }
});

// --- FUNCIÓN DEL PANEL ---
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
        { tag: "ccNombreDoc",     valor: vNombreDoc },
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

// --- BOTONES DE ACCIÓN ---

async function limpiarFormato(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      context.load(selection, "font");
      await context.sync();

      selection.font.set({
        name: "Arial",
        size: 11,
        color: "#000000",
        bold: false,
        italic: false
      });
      await context.sync();
      
      // Intentamos justificar sin romper nada
      context.load(selection, "paragraphFormat");
      await context.sync();
      try { selection.paragraphFormat.alignment = "Justified"; await context.sync(); } catch (e) {}
    });
  } catch (error) { console.error(error); } 
  finally { if (event) event.completed(); }
}

async function insertarFecha(event) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const fechaHoy = new Date().toLocaleDateString();
    selection.insertText(fechaHoy, "Replace");
    await context.sync();
  });
  if (event) event.completed();
}

// --- NUEVOS BOTONES DE ESTILOS (ESPAÑOL) ---

async function estiloTitulo1(event) {
  // En Word en Español, el estilo se llama "Título 1"
  await aplicarEstiloSeguro("Título 1"); 
  if (event) event.completed();
}

async function estiloTitulo2(event) {
  await aplicarEstiloSeguro("Título 2");
  if (event) event.completed();
}

async function estiloTitulo3(event) {
  await aplicarEstiloSeguro("Título 3");
  if (event) event.completed();
}

// Función auxiliar blindada
async function aplicarEstiloSeguro(nombreEstilo) {
  await Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      
      // Aplicamos estilo por nombre local
      selection.style = nombreEstilo;
      
      await context.sync();
    } catch (error) {
      // Si falla (ej: Word en Inglés), intentamos el nombre en inglés
      // Esto es un "Plan B" automático
      try {
        const selection = context.document.getSelection();
        const nombreIngles = nombreEstilo.replace("Título", "Heading");
        selection.style = nombreIngles;
        await context.sync();
      } catch (e2) {
         console.warn("No se encontró el estilo: " + nombreEstilo);
      }
    }
  });
}

// --- REGISTRO ---
Office.actions.associate("limpiarFormato", limpiarFormato);
Office.actions.associate("insertarFecha", insertarFecha);
Office.actions.associate("estiloTitulo1", estiloTitulo1);
Office.actions.associate("estiloTitulo2", estiloTitulo2);
Office.actions.associate("estiloTitulo3", estiloTitulo3);