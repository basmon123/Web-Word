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


// ==========================================
// 2. FUNCIÓN: Limpiar Formato (CORREGIDO)
// ==========================================
async function limpiarFormato(event) {
  await Word.run(async (context) => {
    // CORRECCIÓN CLAVE: Usamos getSelection() en vez de body.
    // getSelection = Solo lo que sombreaste con el mouse.
    const range = context.document.getSelection();
    
    // Cargamos el texto para asegurar que Word procese el rango
    context.load(range, 'text'); 
    await context.sync();

    // Si no hay nada seleccionado, avisamos (opcional) o no hacemos nada
    if (range.text === "") {
        console.log("Nada seleccionado.");
    } else {
        // Aplicamos formato SOLO a la selección
        range.font.name = "Arial";
        range.font.size = 10;
        range.font.color = "black";
        range.font.bold = false;
        range.font.italic = false;
        
        // Limpiamos resaltados
        range.font.highlightColor = null; 

        // Intentamos justificar (protegido contra errores en tablas)
        try {
            range.paragraphFormat.alignment = "Justified";
        } catch (e) {
            console.warn("No se pudo justificar (quizás es tabla).");
        }
    }
    
    await context.sync();
  });
  
  // Avisar a Office que terminamos
  if (event) event.completed();
}

// ==========================================
// 3. FUNCIÓN: Insertar Fecha
// ==========================================
async function insertarFecha(event) {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    const hoy = new Date();
    const opciones = { year: 'numeric', month: 'long', day: 'numeric' };
    const fechaTexto = hoy.toLocaleDateString('es-ES', opciones);
    
    range.insertText(fechaTexto, "Replace");
    await context.sync();
  });
  
  if (event) event.completed();
}

// ==========================================
// 4. FUNCIONES: Estilos de Título (CORREGIDO)
// ==========================================

// Helper: Intenta aplicar estilo en Inglés, si falla, usa Español
async function aplicarEstiloSeguro(nombreIngles, nombreEspanol) {
    await Word.run(async (context) => {
        const range = context.document.getSelection();
        try {
            // Intento 1: Word interno (generalmente Inglés)
            range.style = nombreIngles;
            await context.sync();
        } catch (error) {
            // Intento 2: Word local (Español)
            try {
                range.style = nombreEspanol;
                await context.sync();
            } catch (e2) {
                console.warn("No existe el estilo: " + nombreEspanol);
            }
        }
    });
}

async function estiloTitulo1(event) {
  await aplicarEstiloSeguro("Heading 1", "Título 1");
  if (event) event.completed();
}

async function estiloTitulo2(event) {
  await aplicarEstiloSeguro("Heading 2", "Título 2");
  if (event) event.completed();
}

async function estiloTitulo3(event) {
  await aplicarEstiloSeguro("Heading 3", "Título 3");
  if (event) event.completed();
}

// ==========================================
// 5. REGISTRO DE FUNCIONES (OBLIGATORIO)
// ==========================================

// Mapeo para Office
Office.actions.associate("abrirCatalogo", abrirCatalogo);
Office.actions.associate("limpiarFormato", limpiarFormato);
Office.actions.associate("insertarFecha", insertarFecha);
Office.actions.associate("estiloTitulo1", estiloTitulo1);
Office.actions.associate("estiloTitulo2", estiloTitulo2);
Office.actions.associate("estiloTitulo3", estiloTitulo3);

// Mapeo global (seguridad extra)
const g = typeof globalThis !== "undefined" ? globalThis : window;
g.abrirCatalogo = abrirCatalogo;
g.limpiarFormato = limpiarFormato;
g.insertarFecha = insertarFecha;
g.estiloTitulo1 = estiloTitulo1;
g.estiloTitulo2 = estiloTitulo2;
g.estiloTitulo3 = estiloTitulo3;