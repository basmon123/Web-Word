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

  // URL del archivo en la nube (GitHub)
  // Ajusta esta ruta: sube desde 'commands' a raiz y entra a 'templates'
  const urlPlantilla = "../../templates/" + nombreArchivo;

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