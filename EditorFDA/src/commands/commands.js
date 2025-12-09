/* global Office, Word */

Office.onReady(() => {
  console.log("Office initialized en commands.js");
});

let dialog; 

// ==========================================
// 1. LÓGICA DEL CATÁLOGO (Nuevo Documento)
// ==========================================

function abrirCatalogo(event) {
  // Nota: ?v=4 para asegurar que no cachee viejo
  const url = "https://basmon123.github.io/Web-Word/EditorFDA/src/catalog/catalog.html?v=4";

  Office.context.ui.displayDialogAsync(url, { height: 60, width: 50 },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error.message);
      } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, procesarMensaje);
      }
    }
  );
  
  if(event) event.completed();
}

async function procesarMensaje(arg) {
  dialog.close(); 
  const mensaje = JSON.parse(arg.message); 

  if (mensaje.accion === "CREAR_DOCUMENTO") {
      await crearDocumentoNuevo(mensaje.plantilla, mensaje.datos);
  }
}

// --- FUNCIÓN PRINCIPAL CORREGIDA (MÉTODO DE INYECCIÓN) ---
async function crearDocumentoNuevo(nombrePlantilla, datosProyecto) {
  
  // 1. Mapeo de archivos
  const archivos = {
      "Minuta": "Minuta.docx",
      "Informe": "Informe.docx",
      "Carta": "Carta.docx"
  };

  const nombreArchivo = archivos[nombrePlantilla];
  if (!nombreArchivo) return;

  // Valor por defecto para que no falle la descarga mientras investigamos
  const carpeta = datosProyecto.CarpetaPlantilla || "CODELCO"; 
  const urlPlantilla = "https://basmon123.github.io/Web-Word/EditorFDA/src/templates/" + carpeta + "/" + nombreArchivo;

  try {
      const response = await fetch(urlPlantilla);
      if (!response.ok) throw new Error("Error descargando plantilla");
      
      const blob = await response.blob();
      const base64 = await getBase64FromBlob(blob);

      await Word.run(async (context) => {
        const newDoc = context.application.createDocument(base64);

        // ============================================================
        // 🕵️‍♂️ MODO SUPER ESPÍA: IMPRIMIR EL JSON EN EL DOC
        // ============================================================
        
        // Convertimos TODO el objeto de datos a texto para leerlo
        const datosCrudos = JSON.stringify(datosProyecto, null, 2);
        
        // Lo escribimos en rojo y grande al principio del documento
        const parrafoDebug = newDoc.body.insertParagraph("DATOS RECIBIDOS:\n" + datosCrudos, "Start");
        parrafoDebug.font.color = "red";
        parrafoDebug.font.name = "Courier New"; // Fuente tipo código
        parrafoDebug.font.size = 10;
        
        // Insertamos un salto de página para separar el reporte de la plantilla real
        parrafoDebug.insertBreak(Word.BreakType.page, "After");

        // ============================================================

        // Intento de rellenado (seguramente fallará, pero lo dejamos)
        const mapaDatos = [
            { tag: "ccCliente",    valor: datosProyecto.Cliente },
            { tag: "ccDivisión",   valor: datosProyecto.Division },
            { tag: "ccProyecto",   valor: datosProyecto.NombreProyecto }, 
            { tag: "ccContrato",   valor: datosProyecto.Contrato },
            { tag: "ccAPI",        valor: datosProyecto.API },
            { tag: "ccCodigo",     valor: datosProyecto.Título || datosProyecto.Title }
        ];

        for (let item of mapaDatos) {
            if (!item.valor) continue;
            const controls = newDoc.body.contentControls.getByTag(item.tag);
            controls.load("items");
            await context.sync();
            if (controls.items.length > 0) {
                controls.items.forEach(c => c.insertText(String(item.valor), "Replace"));
            }
        }
        
        newDoc.open();
        await context.sync();
      });

  } catch (error) {
      console.error("ERROR:", error);
  }
}

function getBase64FromBlob(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
            const base64String = reader.result.toString().split(',')[1];
            resolve(base64String);
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
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

// 3. REGISTRO OFICIAL (LA PARTE CLAVE)
// Aquí registramos AMBAS funciones usando el MISMO método.
// Esto elimina la interferencia.
Office.actions.associate("limpiarFormato", limpiarFormato);
Office.actions.associate("insertarFecha", insertarFecha);
Office.actions.associate("estiloTitulo1", estiloTitulo1);
Office.actions.associate("estiloTitulo2", estiloTitulo2);
Office.actions.associate("estiloTitulo3", estiloTitulo3);
Office.actions.associate("abrirCatalogo", abrirCatalogo);
