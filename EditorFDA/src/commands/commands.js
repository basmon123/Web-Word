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

// --- FUNCIÓN PRINCIPAL CORREGIDA (NOMBRES CORRECTOS) ---
async function crearDocumentoNuevo(nombrePlantilla, datosProyecto) {
  
  // 1. Mapeo de archivos
  const archivos = {
      "Minuta": "Minuta.docx",
      "Informe": "Informe.docx",
      "Carta": "Carta.docx"
  };

  const nombreArchivo = archivos[nombrePlantilla];
  if (!nombreArchivo) return;

  // Ajuste: El JSON dice "carpeta_plantilla" (minúscula), no "CarpetaPlantilla"
  const carpeta = datosProyecto.carpeta_plantilla || "CODELCO"; 
  const urlPlantilla = "https://basmon123.github.io/Web-Word/EditorFDA/src/templates/" + carpeta + "/" + nombreArchivo;

  try {
      // 2. Descargar la plantilla
      const response = await fetch(urlPlantilla);
      if (!response.ok) throw new Error("No se encontró la plantilla");
      
      const blob = await response.blob();
      const base64 = await getBase64FromBlob(blob);

      await Word.run(async (context) => {
        // 3. Crear el documento en memoria (NO ABRIR AÚN)
        const newDoc = context.application.createDocument(base64);

        // --- 4. RELLENADO DE DATOS (AHORA SÍ COINCIDEN) ---
        // Izquierda (tag): Etiqueta en Word
        // Derecha (valor): Propiedad EXACTA de tu JSON
        const mapaDatos = [
            { tag: "ccCliente",    valor: datosProyecto.cliente },
            { tag: "ccDivisión",   valor: datosProyecto.division },
            { tag: "ccProyecto",   valor: datosProyecto.nombre },  // Antes era 'NombreProyecto'
            { tag: "ccContrato",   valor: datosProyecto.contrato },
            { tag: "ccAPI",        valor: datosProyecto.api },
            { tag: "ccID",     valor: datosProyecto.id }       // El ID es el número (7560)
        ];

        for (let item of mapaDatos) {
            // Si el dato viene vacío, saltamos al siguiente
            if (!item.valor) continue;

            // Buscamos la cajita en el NUEVO documento
            const controls = newDoc.body.contentControls.getByTag(item.tag);
            controls.load("items");
            
            await context.sync();

            // Si la encontramos, escribimos dentro
            if (controls.items.length > 0) {
                controls.items.forEach((control) => {
                    // Convertimos a String por seguridad
                    control.insertText(String(item.valor), "Replace");
                });
            }
        }

        // 5. Abrimos el documento ya rellenado
        newDoc.open();
        await context.sync();
        // =======================================================
        // 6. NUEVO: CERRAR EL DOCUMENTO ACTUAL (EL EN BLANCO)
        // =======================================================
        // 'context.document' se refiere al documento desde donde lanzaste el comando.
        // 'skipSave' evita que pregunte "¿Desea guardar?" al cerrarse.
        context.document.close(Word.CloseBehavior.skipSave); 
        // =======================================================
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

