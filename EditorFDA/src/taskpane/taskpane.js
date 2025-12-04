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

// --- CÓDIGO AÑADIDO PARA EL BOTÓN DE FECHA ---

function insertarFecha(event) {
  Word.run(function (context) {
    
    // 1. Insertar fecha
    var body = context.document.body;
    var hoy = new Date().toLocaleDateString();
    
    // Insertamos al inicio del documento solo para probar que está vivo
    body.insertParagraph("FECHA DESDE TASKPANE.JS: " + hoy, "Start");

    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
  })
  .finally(function () {
      // 2. Avisar que terminó (CRÍTICO)
      if (event) {
          event.completed();
      }
  });
}

// Registro de la función en el mismo archivo del panel
// NOTA: Asegúrate de que "insertarFecha" sea igual al <FunctionName> en tu XML
Office.actions.associate("insertarFecha", insertarFecha);