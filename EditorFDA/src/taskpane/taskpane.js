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

      context.load(selection, "paragraphFormat");
      await context.sync();
      try {
        selection.paragraphFormat.alignment = "Justified";
        await context.sync();
      } catch (e) { console.warn("Alineación omitida"); }
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

// --- NUEVOS BOTONES DE ESTILOS ---
// Estos aplican los estilos internos de Word "Heading 1", "Heading 2", etc.
// En Word en español, esto corresponde a "Título 1", "Título 2"...

async function estiloTitulo1(event) {
  await aplicarEstilo("Heading 1");
  if (event) event.completed();
}

async function estiloTitulo2(event) {
  await aplicarEstilo("Heading 2");
  if (event) event.completed();
}

async function estiloTitulo3(event) {
  await aplicarEstilo("Heading 3");
  if (event) event.completed();
}

// Función auxiliar para no repetir código
async function aplicarEstilo(nombreEstilo) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    // Aplicamos el estilo al párrafo seleccionado
    selection.style = nombreEstilo; 
    await context.sync();
  });
}

// --- REGISTRO DE TODAS LAS FUNCIONES ---
Office.actions.associate("limpiarFormato", limpiarFormato);
Office.actions.associate("insertarFecha", insertarFecha);
Office.actions.associate("estiloTitulo1", estiloTitulo1);
Office.actions.associate("estiloTitulo2", estiloTitulo2);
Office.actions.associate("estiloTitulo3", estiloTitulo3);