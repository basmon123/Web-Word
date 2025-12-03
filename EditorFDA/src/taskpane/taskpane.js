/* global document, Office, Word */



Office.onReady((info) => {

  if (info.host === Office.HostType.Word) {

    const btn = document.getElementById("btnGenerar");

    if (btn) btn.onclick = run;

  }

});



// --- 1. FUNCIÓN DEL PANEL (GENERADOR) ---

async function run() {

  try {

    const getVal = (id) => document.getElementById(id) ? document.getElementById(id).value : "";

    const msgLabel = document.getElementById("mensajeEstado");

   

    // Captura rápida

    const datos = {

        "ccCliente": getVal("inCliente"),

        "ccDivisión": getVal("inDivision"),

        "ccServicios": getVal("inServicios"),

        "ccContrato": getVal("inContrato"),

        "ccAPI": getVal("inAPI"),

        "ccProyecto": getVal("inProyecto"),

        "ccNombreDoc": getVal("inNombreDoc"),

        "ccCodigo": getVal("inCodigo"),

        "ccRevision": getVal("inRevision")

    };



    if (msgLabel) msgLabel.textContent = "Procesando...";



    await Word.run(async (context) => {

      let contadores = 0;

     

      // Lógica unificada para cuerpo y encabezados

      const tagsMapa = [

          { t: "ccCliente", v: datos.ccCliente }, { t: "ccCliente_encabezado", v: datos.ccCliente },

          { t: "ccDivisión", v: datos.ccDivisión }, { t: "ccD_encabezado", v: datos.ccDivisión },

          { t: "ccServicios", v: datos.ccServicios },

          { t: "ccContrato", v: datos.ccContrato },

          { t: "ccAPI", v: datos.ccAPI },

          { t: "ccProyecto", v: datos.ccProyecto }, { t: "ccNProyecto_Encabezado", v: datos.ccProyecto },

          { t: "ccNombreDoc", v: datos.ccNombreDoc }, { t: "ccNombre doc", v: datos.ccNombreDoc }, // Por si acaso

          { t: "ccCodigo", v: datos.ccCodigo },

          { t: "ccRevision", v: datos.ccRevision }

      ];



      for (let item of tagsMapa) {

        if(!item.v) continue; // Si está vacío, saltar

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



// --- 2. HERRAMIENTAS DE FORMATO ---



async function limpiarFormato(event) {

  try {

    await Word.run(async (context) => {

      const selection = context.document.getSelection();

     

      // Paso 1: Fuente

      context.load(selection, "font");

      await context.sync();

      selection.font.set({ name: "Arial", size: 11, color: "#000000", bold: false, italic: false });

      await context.sync();

     

      // Paso 2: Párrafo (Intento seguro)

      context.load(selection, "paragraphFormat");

      await context.sync();

      try {

          selection.paragraphFormat.alignment = "Justified";

          await context.sync();

      } catch (e) {

          console.warn("No se pudo justificar (posible tabla o restricción).");

      }

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



// --- 3. ESTILOS FDA (1.0, 1.1, 1.1.1) ---



async function estiloTitulo1(event) {

  await aplicarEstiloProfesional("Título 1", "Heading 1");

  if (event) event.completed();

}



async function estiloTitulo2(event) {

  await aplicarEstiloProfesional("Título 2", "Heading 2");

  if (event) event.completed();

}



async function estiloTitulo3(event) {

  await aplicarEstiloProfesional("Título 3", "Heading 3");

  if (event) event.completed();

}



// Función auxiliar inteligente (Prueba Español -> Falla -> Prueba Inglés)

async function aplicarEstiloProfesional(nombreEsp, nombreIng) {

  await Word.run(async (context) => {

    try {

      const selection = context.document.getSelection();

      selection.style = nombreEsp; // Intento Español

      await context.sync();

    } catch (error) {

      // Si falla, intentamos Inglés silenciosamente

      try {

        const selection = context.document.getSelection();

        selection.style = nombreIng;

        await context.sync();

      } catch (e2) {

        console.warn("No se encontró el estilo ni en ESP ni ING.");

      }

    }

  });

}



// --- 4. REGISTRO DE FUNCIONES ---

Office.actions.associate("limpiarFormato", limpiarFormato);

Office.actions.associate("insertarFecha", insertarFecha);

Office.actions.associate("estiloTitulo1", estiloTitulo1);

Office.actions.associate("estiloTitulo2", estiloTitulo2);

Office.actions.associate("estiloTitulo3", estiloTitulo3);