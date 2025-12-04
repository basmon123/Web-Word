/* global Office, Word */

Office.onReady(() => {
  // Inicialización lista
  console.log("Office initialized en commands.js");
});

let dialog; 

// ==========================================
// 1. LÓGICA DEL CATÁLOGO (Nuevo Documento)
// ==========================================

function abrirCatalogo(event) {
  // Nota: ?v=3 para asegurar que no cachee viejo
  const url = "https://basmon123.github.io/Web-Word/EditorFDA/src/catalog/catalog.html?v=3";

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
  
  // Avisamos a la cinta que el botón se presionó correctamente
  if(event) event.completed();
}

async function procesarMensaje(arg) {
  dialog.close(); 
  const mensaje = JSON.parse(arg.message); 

  if (mensaje.accion === "CREAR_DOCUMENTO") {
      await crearDocumentoNuevo(mensaje.plantilla, mensaje.datos);
  }
}

async function crearDocumentoNuevo(nombrePlantilla, datosProyecto) {
  const archivos = {
      "Minuta": "Minuta.docx",
      "Informe": "Informe.docx",
      "Carta": "Carta.docx"
  };

  const nombreArchivo = archivos[nombrePlantilla];
  if (!nombreArchivo) return;

  const urlPlantilla = "https://basmon123.github.io/Web-Word/EditorFDA/src/templates/" + datosProyecto.carpeta_plantilla + "/" + nombreArchivo;

  try {
      const response = await fetch(urlPlantilla);
      if (!response.ok) throw new Error("No se encontró la plantilla");
      
      const blob = await response.blob();
      const base64 = await getBase64FromBlob(blob);

      await Word.run(async (context) => {
        const newDoc = context.application.createDocument(base64);
        newDoc.open();
        await context.sync();
      });

  } catch (error) {
      console.error("Error al crear documento:", error);
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

// ==========================================
// 2. LÓGICA DE HERRAMIENTAS (Limpiar Formato)
// ==========================================

// --- FUNCIÓN DE FECHA ---
async function insertarFecha(event) {
  try {
    await Word.run(async (context) => {
      // 1. Obtener la ubicación del cursor
      const range = context.document.getSelection();
      
      // 2. Obtener fecha de hoy
      const hoy = new Date();
      const fechaTexto = hoy.toLocaleDateString("es-ES", {
          year: 'numeric', month: 'long', day: 'numeric'
      });
      
      // 3. Escribir en Word
      range.insertText(fechaTexto, "Replace");
      
      // 4. Sincronizar
      await context.sync();
    });
  } catch (error) {
    // Si falla, escribimos el error en el documento para verlo
    await Word.run(async (context) => {
        context.document.body.insertParagraph("ERROR JS: " + error.message, "Start");
        await context.sync();
    });
  } finally {
    // 5. IMPORTANTE: Avisar al botón que terminó
    if (event) {
        event.completed();
    }
  }
}


// ==========================================
// 3. REGISTRO OFICIAL (LA PARTE CLAVE)
// ==========================================
// Aquí registramos AMBAS funciones usando el MISMO método.
// Esto elimina la interferencia.

Office.actions.associate("abrirCatalogo", abrirCatalogo);
Office.actions.associate("insertarFecha", insertarFecha);