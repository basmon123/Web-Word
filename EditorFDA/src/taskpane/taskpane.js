/* global document, Office, Word */

// --- 1. BASE DE DATOS PROYECTOS (Simulación) ---
const dbProyectos = [
  { id: "7560", cliente: "CODELCO CHILE", division: "Gabriela Mistral", nombre: "Estudio Diagnóstico Pila ROM", contrato: "4600025605", api: "G25D203" },
  { id: "8890", cliente: "ANGLO AMERICAN", division: "Los Bronces", nombre: "Ingeniería de Detalles Tranque", contrato: "550001234", api: "A99X100" },
  { id: "1020", cliente: "BHP", division: "Escondida", nombre: "Optimización Sistema Bombeo", contrato: "330009876", api: "B50Z200" }
];

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Botón Principal (Generar)
    const btnGen = document.getElementById("btnGenerar");
    if (btnGen) btnGen.onclick = run;

    // Nuevo Botón (Buscar)
    const btnBus = document.getElementById("btnBuscar");
    if (btnBus) btnBus.onclick = buscarProyecto;
  }
});

// --- 2. FUNCIÓN DE BÚSQUEDA (ACTUALIZADA) ---
function buscarProyecto() {
    const idBusqueda = document.getElementById("inputBusqueda").value;
    const msg = document.getElementById("msgBusqueda");

    // Limpiamos clases previas
    msg.className = "mensaje-busqueda"; 

    // Buscamos en la "Base de Datos"
    const proyecto = dbProyectos.find(p => p.id === idBusqueda);

    if (proyecto) {
        msg.textContent = "✅ Proyecto encontrado. Datos cargados.";
        msg.classList.add("texto-exito"); // Añade clase verde del CSS

        // AUTO-COMPLETAMOS
        document.getElementById("inCliente").value = proyecto.cliente;
        document.getElementById("inDivision").value = proyecto.division;
        document.getElementById("inProyecto").value = proyecto.nombre;
        document.getElementById("inContrato").value = proyecto.contrato;
        document.getElementById("inAPI").value = proyecto.api;
        
        document.getElementById("inNombreDoc").value = "";
        document.getElementById("inCodigo").value = "";
        
    } else {
        msg.textContent = "❌ Proyecto no encontrado.";
        msg.classList.add("texto-error"); // Añade clase roja del CSS
    }
}

// --- 3. FUNCIÓN DEL PANEL (GENERADOR) - Esta es la tuya de ayer ---
async function run() {
  try {
    const getVal = (id) => document.getElementById(id) ? document.getElementById(id).value : "";
    const msgLabel = document.getElementById("mensajeEstado");
    
    // Captura lo que haya en los inputs (sea manual o autocompletado)
    const datos = {
        "ccCliente": getVal("inCliente"),
        "ccDivisión": getVal("inDivision"),
        "ccServicios": getVal("inServicios"),
        "ccContrato": getVal("inContrato"),
        "ccAPI": getVal("inAPI"),
        "ccProyecto": getVal("inProyecto"),
        "ccNombreDoc": getVal("inNombreDoc"),
        "ccNombre doc": getVal("inNombreDoc"),
        "ccCodigo": getVal("inCodigo"),
        "ccRevision": getVal("inRevision"),
        
        // Encabezados (Repetimos datos)
        "ccCliente_encabezado": getVal("inCliente"),
        "ccD_encabezado": getVal("inDivision"),
        "ccNProyecto_Encabezado": getVal("inProyecto")
    };

    if (msgLabel) msgLabel.textContent = "Procesando...";

    await Word.run(async (context) => {
      let contadores = 0;
      
      // Lista unificada de tags
      const tagsMapa = [
          { t: "ccCliente", v: datos.ccCliente }, 
          { t: "ccDivisión", v: datos.ccDivisión }, 
          { t: "ccServicios", v: datos.ccServicios },
          { t: "ccContrato", v: datos.ccContrato },
          { t: "ccAPI", v: datos.ccAPI },
          { t: "ccProyecto", v: datos.ccProyecto }, 
          { t: "ccNombreDoc", v: datos.ccNombreDoc }, 
          { t: "ccNombre doc", v: datos.ccNombreDoc },
          { t: "ccCodigo", v: datos.ccCodigo },
          { t: "ccRevision", v: datos.ccRevision },
          // Encabezados
          { t: "ccCliente_encabezado", v: datos.ccCliente },
          { t: "ccD_encabezado", v: datos.ccDivisión },
          { t: "ccNProyecto_Encabezado", v: datos.ccProyecto }
      ];

      for (let item of tagsMapa) {
        if(!item.v) continue; 
        let ccs = context.document.contentControls.getByTag(item.t);
        ccs.load("items");
        await context.sync();
        if (ccs.items.length > 0) {
           for (let cc of ccs.items) {
             cc.insertText(item.v, "Replace");
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

// --- 4. HERRAMIENTAS DE LA CINTA (NO TOCAR - MANTENER) ---
// Estas son las que usa tu barra de herramientas personalizada

async function limpiarFormato(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      context.load(selection, "font");
      await context.sync();
      selection.font.set({ name: "Arial", size: 11, color: "#000000", bold: false, italic: false });
      await context.sync();
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

// Estilos
async function estiloTitulo1(event) { await aplicarEstiloProfesional("Título 1", "Heading 1"); if (event) event.completed(); }
async function estiloTitulo2(event) { await aplicarEstiloProfesional("Título 2", "Heading 2"); if (event) event.completed(); }
async function estiloTitulo3(event) { await aplicarEstiloProfesional("Título 3", "Heading 3"); if (event) event.completed(); }

async function aplicarEstiloProfesional(nombreEsp, nombreIng) {
  await Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      selection.style = nombreEsp; 
      await context.sync();
    } catch (error) {
      try {
        const selection = context.document.getSelection();
        selection.style = nombreIng;
        await context.sync();
      } catch (e2) {}
    }
  });
}

// REGISTRO
Office.actions.associate("limpiarFormato", limpiarFormato);
Office.actions.associate("insertarFecha", insertarFecha);
Office.actions.associate("estiloTitulo1", estiloTitulo1);
Office.actions.associate("estiloTitulo2", estiloTitulo2);
Office.actions.associate("estiloTitulo3", estiloTitulo3);