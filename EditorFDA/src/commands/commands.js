/* global Office, Word */

// Variables globales para los diálogos
let dialogCatalogo; 
let dialogGenerador;

Office.onReady(() => {
  console.log("Office initialized en commands.js");
});

// ==========================================
// 1. LÓGICA DEL CATÁLOGO (Nuevo Documento)
// ==========================================

function abrirCatalogo(event) {
  const url = "https://basmon123.github.io/Web-Word/EditorFDA/src/catalog/catalog.html?v=4";

  Office.context.ui.displayDialogAsync(url, { height: 60, width: 50 },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error.message);
      } else {
        dialogCatalogo = asyncResult.value;
        dialogCatalogo.addEventHandler(Office.EventType.DialogMessageReceived, procesarMensajeCatalogo);
      }
    }
  );
  
  if(event) event.completed();
}

async function procesarMensajeCatalogo(arg) {
  dialogCatalogo.close(); 
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

  const carpeta = datosProyecto.carpeta_plantilla || "CODELCO"; 
  const urlPlantilla = "https://basmon123.github.io/Web-Word/EditorFDA/src/templates/" + carpeta + "/" + nombreArchivo;

  try {
      const response = await fetch(urlPlantilla);
      if (!response.ok) throw new Error("No se encontró la plantilla");
      
      const blob = await response.blob();
      const base64 = await getBase64FromBlob(blob);

      await Word.run(async (context) => {
        const newDoc = context.application.createDocument(base64);

        const mapaDatos = [
            { tag: "ccCliente",    valor: datosProyecto.cliente },
            { tag: "ccDivisión",   valor: datosProyecto.division },
            { tag: "ccProyecto",   valor: datosProyecto.nombre },
            { tag: "ccContrato",   valor: datosProyecto.contrato },
            { tag: "ccAPI",        valor: datosProyecto.api },
            { tag: "ccID",         valor: datosProyecto.id }
        ];

        for (let item of mapaDatos) {
            if (!item.valor) continue;
            const controls = newDoc.body.contentControls.getByTag(item.tag);
            controls.load("items");
            await context.sync();
            if (controls.items.length > 0) {
                controls.items.forEach((control) => {
                    control.insertText(String(item.valor), "Replace");
                });
            }
        }

        newDoc.open();
        await context.sync();
        context.document.close(Word.CloseBehavior.skipSave); 
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
// 2. HERRAMIENTAS DE FORMATO
// ==========================================

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
      try { 
          selection.paragraphFormat.alignment = "Justified"; 
          await context.sync(); 
      } catch (e) { console.warn("No se pudo justificar."); }
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

// ==========================================
// 3. ESTILOS FDA
// ==========================================

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
      } catch (e2) { console.warn("Estilo no encontrado."); }
    }
  });
}

// ==========================================
// 4. NUEVO: GENERADOR DE TABLAS (CORREGIDO)
// ==========================================

function abrirVentanaTablas(event) {
    const url = "https://basmon123.github.io/Web-Word/EditorFDA/src/GeneradorTablas/generadorTablas.html"; 

    const opciones = { height: 45, width: 30, displayInIframe: true };

    Office.context.ui.displayDialogAsync(url, opciones, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Error al abrir diálogo:", asyncResult.error.message);
        } else {
            dialogGenerador = asyncResult.value;
            dialogGenerador.addEventHandler(Office.EventType.DialogMessageReceived, procesarMensajeTabla);
        }
    });
    
    if(event) event.completed();
}

// --- ESTA ES LA FUNCIÓN QUE ARREGLA TODO ---
async function procesarMensajeTabla(arg) {
    let datos;
    try { datos = JSON.parse(arg.message); } catch (e) { return; }
    
    // IMPORTANTE: Solo cerramos la ventana si NO estamos escaneando.
    // Si estamos escaneando, queremos que la ventana siga abierta para copiar el código.
    if (datos.accion !== "EXTRAER_XML") {
        if (dialogGenerador) dialogGenerador.close();
    }

    await Word.run(async (context) => {
        
        // CASO 1: INSERTAR TABLA SIMPLE (MANUAL)
        if (datos.accion === "INSERTAR") {
            const seleccion = context.document.getSelection();
            
            let matriz = [];
            const filas = parseInt(datos.filas);
            const cols = parseInt(datos.columnas);

            for(let i=0; i<filas; i++) {
                let fila = new Array(cols).fill(" "); 
                matriz.push(fila);
            }

            const tabla = seleccion.insertTable(filas, cols, "After", matriz);
            tabla.autofitWindow();
            await context.sync();
        } 
        
        // CASO 2: INSERTAR PLANTILLA DESDE JSON (BIBLIOTECA)
        else if (datos.accion === "INSERTAR_XML") {
            const seleccion = context.document.getSelection();
            
            // Insertamos el ADN de la tabla guardada
            seleccion.insertOoxml(datos.xml, "After");
            
            // Un salto de línea para separar
            seleccion.insertParagraph("", "After");
            
            await context.sync();
        }

        // CASO 3: ESCANEAR TABLA (HERRAMIENTA DESARROLLADOR)
        else if (datos.accion === "EXTRAER_XML") {
            const seleccion = context.document.getSelection();
            
            // 1. Obtenemos el código de la tabla seleccionada
            const xml = seleccion.getOoxml();
            await context.sync();
            
            // 2. Escribimos el código TEMPORALMENTE en la hoja de Word
            // (Es la forma más rápida para que lo puedas copiar)
            seleccion.insertText(xml.value, "Replace");
            await context.sync();
        }

    }).catch(error => console.error("Error en tabla:", error));
}


// ==========================================
// 5. REGISTRO OFICIAL
// ==========================================

Office.actions.associate("limpiarFormato", limpiarFormato);
Office.actions.associate("insertarFecha", insertarFecha);
Office.actions.associate("estiloTitulo1", estiloTitulo1);
Office.actions.associate("estiloTitulo2", estiloTitulo2);
Office.actions.associate("estiloTitulo3", estiloTitulo3);
Office.actions.associate("abrirCatalogo", abrirCatalogo);
Office.actions.associate("abrirVentanaTablas", abrirVentanaTablas);