/* global Office, Word */

let dialogCatalogo; 
let dialogGenerador;

Office.onReady(() => {
  console.log("Commands.js listo");
});

// ==========================================
// 1. LÓGICA DE TABLAS
// ==========================================

function abrirVentanaTablas(event) {
    const url = "https://basmon123.github.io/Web-Word/EditorFDA/src/GeneradorTablas/generadorTablas.html"; 
    // Asegúrate de que displayInIframe sea true para mejor comunicación
    const opciones = { height: 50, width: 30, displayInIframe: true };

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

// REEMPLAZA ESTA FUNCIÓN EN COMMANDS.JS
async function procesarMensajeTabla(arg) {
    let datos;
    try { datos = JSON.parse(arg.message); } catch (e) { return; }
    
    // No cerrar ventana si estamos escaneando
    if (datos.accion !== "EXTRAER_XML") {
        if (dialogGenerador) dialogGenerador.close();
    }

    await Word.run(async (context) => {
        
        // CASO A: INSERTAR MANUAL
        if (datos.accion === "INSERTAR") {
            const seleccion = context.document.getSelection();
            let matriz = [];
            for(let i=0; i<parseInt(datos.filas); i++) {
                let fila = new Array(parseInt(datos.columnas)).fill(" "); 
                matriz.push(fila);
            }
            const tabla = seleccion.insertTable(parseInt(datos.filas), parseInt(datos.columnas), "After", matriz);
            tabla.autofitWindow();
            await context.sync();
        } 
        
        // CASO B: INSERTAR DESDE BIBLIOTECA (CON DIAGNÓSTICO)
        else if (datos.accion === "INSERTAR_XML") {
            const seleccion = context.document.getSelection();
            
            // INTENTO DE INSERCIÓN
            try {
                // Insertamos el código
                seleccion.insertOoxml(datos.xml, "After");
                seleccion.insertParagraph("", "After");
                await context.sync();
            } catch (errorXML) {
                // SI FALLA, AVISAMOS AL USUARIO
                console.error("XML Corrupto:", errorXML);
                
                // Escribimos el error en el documento para que lo veas
                const body = context.document.body;
                body.insertParagraph("❌ ERROR: El código XML de esta tabla está dañado o incompleto.", "Start");
                body.insertParagraph("Detalle: " + errorXML.message, "Start");
                await context.sync();
            }
        }

        // CASO C: ESCANEAR
        else if (datos.accion === "EXTRAER_XML") {
            const seleccion = context.document.getSelection();
            const xmlResult = seleccion.getOoxml();
            await context.sync();
            
            const body = context.document.body;
            body.insertParagraph("--- INICIO CÓDIGO XML ---", "End");
            body.insertParagraph(xmlResult.value, "End");
            body.insertParagraph("--- FIN CÓDIGO XML ---", "End");
            await context.sync();
        }

    }).catch(error => {
        console.error("Error general:", error);
    });
}


// ==========================================
// 2. LÓGICA DEL CATÁLOGO (ANTERIOR)
// ==========================================
function abrirCatalogo(event) {
  const url = "https://basmon123.github.io/Web-Word/EditorFDA/src/catalog/catalog.html?v=4";
  Office.context.ui.displayDialogAsync(url, { height: 60, width: 50 },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) { console.error(asyncResult.error.message);
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
  const archivos = { "Minuta": "Minuta.docx", "Informe": "Informe.docx", "Carta": "Carta.docx" };
  const nombreArchivo = archivos[nombrePlantilla];
  if (!nombreArchivo) return;
  const carpeta = datosProyecto.carpeta_plantilla || "CODELCO"; 
  const urlPlantilla = "https://basmon123.github.io/Web-Word/EditorFDA/src/templates/" + carpeta + "/" + nombreArchivo;

  try {
      const response = await fetch(urlPlantilla);
      if (!response.ok) throw new Error("Plantilla no encontrada");
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
                controls.items.forEach((control) => control.insertText(String(item.valor), "Replace"));
            }
        }
        newDoc.open();
        await context.sync();
        context.document.close(Word.CloseBehavior.skipSave); 
      });
  } catch (error) { console.error(error); }
}

function getBase64FromBlob(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result.toString().split(',')[1]);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

// ==========================================
// 3. HERRAMIENTAS Y ESTILOS
// ==========================================
async function limpiarFormato(event) {
  await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.font.set({ name: "Arial", size: 11, color: "#000000", bold: false, italic: false });
      selection.paragraphFormat.alignment = "Justified";
      await context.sync();
  });
  if (event) event.completed();
}

async function insertarFecha(event) {
  await Word.run(async (context) => {
    context.document.getSelection().insertText(new Date().toLocaleDateString(), "Replace");
    await context.sync();
  });
  if (event) event.completed();
}

async function estiloTitulo1(event) { await aplicarEstilo("Título 1", "Heading 1"); if (event) event.completed(); }
async function estiloTitulo2(event) { await aplicarEstilo("Título 2", "Heading 2"); if (event) event.completed(); }
async function estiloTitulo3(event) { await aplicarEstilo("Título 3", "Heading 3"); if (event) event.completed(); }

async function aplicarEstilo(nomEsp, nomIng) {
  await Word.run(async (context) => {
    try {
      context.document.getSelection().style = nomEsp;
      await context.sync();
    } catch (e) {
      try {
        context.document.getSelection().style = nomIng;
        await context.sync();
      } catch (e2) {}
    }
  });
}

// ==========================================
// 4. REGISTRO OFICIAL
// ==========================================
Office.actions.associate("limpiarFormato", limpiarFormato);
Office.actions.associate("insertarFecha", insertarFecha);
Office.actions.associate("estiloTitulo1", estiloTitulo1);
Office.actions.associate("estiloTitulo2", estiloTitulo2);
Office.actions.associate("estiloTitulo3", estiloTitulo3);
Office.actions.associate("abrirCatalogo", abrirCatalogo);
Office.actions.associate("abrirVentanaTablas", abrirVentanaTablas);