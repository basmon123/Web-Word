/* global Office, Word */



Office.onReady(() => {

  // Inicialización si fuera necesaria

});



let dialog; // Variable para guardar la ventana



// 1. ESTA FUNCIÓN LA LLAMA EL BOTÓN DE LA CINTA

function abrirCatalogo(event) {

  // TRUCO: Agregamos '?v=2' al final para romper el caché de Word

  const url = "https://basmon123.github.io/Web-Word/EditorFDA/src/catalog/catalog.html?v=2";



  // Abrimos la ventana emergente

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





// 2. ESTA FUNCIÓN RECIBE LA ORDEN DE LA VENTANA

async function procesarMensaje(arg) {

  dialog.close(); // Cerramos la ventana primero

 

  const mensaje = JSON.parse(arg.message); // Leemos el JSON {accion, plantilla, datos}



  if (mensaje.accion === "CREAR_DOCUMENTO") {

      await crearDocumentoNuevo(mensaje.plantilla, mensaje.datos);

  }

}



// 3. ESTA FUNCIÓN CREA EL WORD DESDE UNA PLANTILLA REAL

async function crearDocumentoNuevo(nombrePlantilla, datosProyecto) {

 

  // Mapeo: Nombre del icono -> Nombre del archivo real

  const archivos = {

      "Minuta": "Minuta.docx",

      "Informe": "Informe.docx",

      "Carta": "Carta.docx"

  };



  const nombreArchivo = archivos[nombrePlantilla];

  if (!nombreArchivo) return;



// ANTES: .../templates/" + datosProyecto.id + "/"...

// AHORA: Usamos .carpeta_plantilla (Ej: CODELCO)

const urlPlantilla = "https://basmon123.github.io/Web-Word/EditorFDA/src/templates/" + datosProyecto.carpeta_plantilla + "/" + nombreArchivo;



  try {

      // A. DESCARGAR EL ARCHIVO WORD

      const response = await fetch(urlPlantilla);

      if (!response.ok) throw new Error("No se encontró la plantilla");

     

      // B. CONVERTIR A BLOB (Archivo binario)

      const blob = await response.blob();



      // C. CONVERTIR A BASE64 (Lo que entiende Word)

      const base64 = await getBase64FromBlob(blob);



      await Word.run(async (context) => {

        // D. CREAR DOCUMENTO USANDO EL BASE64

        const newDoc = context.application.createDocument(base64);

       

        // E. INYECTAR DATOS DEL PROYECTO (Opcional, pero recomendado)

        // Aquí podrías buscar los Tags en la plantilla nueva y llenarlos

        // Para simplificar, primero abrimos el documento.

       

        newDoc.open();

        await context.sync();

       

        // F. LLENADO DE DATOS (Una vez abierto, en el nuevo contexto)

        // Nota: Esto requiere un manejo de contexto avanzado.

        // Por ahora, logremos que ABRA la plantilla real. El llenado lo hacemos en el paso siguiente.

      });



  } catch (error) {

      console.error("Error al crear documento:", error);

  }

}



// Función auxiliar para convertir archivos a texto base64

function getBase64FromBlob(blob) {

    return new Promise((resolve, reject) => {

        const reader = new FileReader();

        reader.onloadend = () => {

            // El resultado viene como "data:application/vnd.openxml...;base64,....."

            // Word solo quiere la parte después de la coma.

            const base64String = reader.result.toString().split(',')[1];

            resolve(base64String);

        };

        reader.onerror = reject;

        reader.readAsDataURL(blob);

    });

}

// ... (Mantén el registro del final g.abrirCatalogo = ... ) ...

const g = typeof globalThis !== "undefined" ? globalThis : window;

g.abrirCatalogo = abrirCatalogo;


// ====================================================================
//    NUEVAS FUNCIONES AGREGADAS PARA LOS BOTONES DEL MANIFEST
// ====================================================================

// A. Limpiar formato (Arial 10 + Justificado)
/async function limpiarFormato(event) {
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