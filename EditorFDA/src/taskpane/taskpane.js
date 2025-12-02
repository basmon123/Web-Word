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
  await probarEstilo("Título 1", "Heading 1");
  if (event) event.completed();
}

async function estiloTitulo2(event) {
  await probarEstilo("Título 2", "Heading 2");
  if (event) event.completed();
}

async function estiloTitulo3(event) {
  await probarEstilo("Título 3", "Heading 3");
  if (event) event.completed();
}

// Función que escribe en el documento lo que está pasando
async function probarEstilo(nombreEsp, nombreIng) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    
    // 1. AVISO DE INICIO
    selection.insertText(" [INTENTO 1: " + nombreEsp + "] ", "End");
    await context.sync();

    try {
      // Intento 1: Nombre en Español
      selection.style = nombreEsp;
      await context.sync();
      selection.insertText(" [¡EXITO ESPAÑOL!] ", "End");
    } catch (error1) {
      
      // 2. SI FALLA, PROBAMOS INGLÉS
      selection.insertText(" [FALLÓ (" + error1.message + ") -> INTENTO 2: " + nombreIng + "] ", "End");
      await context.sync();
      
      try {
        selection.style = nombreIng;
        await context.sync();
        selection.insertText(" [¡EXITO INGLÉS!] ", "End");
      } catch (error2) {
        // 3. SI FALLA TODO
        selection.insertText(" [ERROR TOTAL: " + error2.message + "] ", "End");
        selection.font.color = "red";
        await context.sync();
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