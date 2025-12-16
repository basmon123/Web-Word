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

/* Reemplaza TU función procesarMensajeTabla actual con ESTA */

async function procesarMensajeTabla(arg) {
    let datos;
    try { datos = JSON.parse(arg.message); } catch (e) { return; }
    
    // 1. GESTIÓN DE LA VENTANA
    // Si es Insertar (Manual o Plantilla), cerramos la ventana para ver el resultado.
    // Si es Escanear, LA DEJAMOS ABIERTA.
    if (datos.accion !== "EXTRAER_XML") {
        if (dialogGenerador) dialogGenerador.close();
    }

    await Word.run(async (context) => {
        
        // ==========================================
        // CASO A: INSERTAR MANUAL (SOLUCIÓN ROBUSTA)
        // ==========================================
        if (datos.accion === "INSERTAR") {
            const seleccion = context.document.getSelection();
            
            const filas = parseInt(datos.filas);
            const cols = parseInt(datos.columnas);
            
            // 1. Construimos la matriz de datos vacíos
            let matriz = [];
            for(let i=0; i<filas; i++) {
                // Llenamos con un espacio para que la celda no colapse
                let fila = new Array(cols).fill(" "); 
                matriz.push(fila);
            }

            // 2. Insertamos la tabla (Esto es lo importante)
            const tabla = seleccion.insertTable(filas, cols, "After", matriz);
            
            // 3. INTENTO DE ESTILO (A prueba de fallos)
            // Envolvemos esto en un try/catch para que si el nombre del estilo no existe
            // en tu idioma, la tabla SE CREE IGUAL (aunque sea fea).
            try {
                // Intenta estilo estándar en Inglés
                tabla.style = "Table Grid"; 
            } catch (errorEstilo) {
                try {
                    // Intento alternativo en Español
                    tabla.style = "Tabla con cuadrícula"; 
                } catch (e2) {
                    // Si todo falla, no hacemos nada y dejamos la tabla sin estilo
                    console.warn("No se pudo aplicar estilo, pero la tabla se creó.");
                }
            }

            tabla.autofitWindow();
            await context.sync();
        } 
        
        // ==========================================
        // CASO B: INSERTAR DESDE BIBLIOTECA (XML)
        // ==========================================
        else if (datos.accion === "INSERTAR_XML") {
            const seleccion = context.document.getSelection();
            try {
                seleccion.insertOoxml(datos.xml, "After");
                seleccion.insertParagraph("", "After"); // Separador
                await context.sync();
            } catch (errorXML) {
                // Si el XML falla, escribimos el error en el documento
                const body = context.document.body;
                body.insertParagraph("❌ ERROR: XML corrupto. " + errorXML.message, "Start");
                await context.sync();
            }
        }

        // ==========================================
        // CASO C: ESCANEAR (HERRAMIENTA DEVELOPER)
        // ==========================================
        else if (datos.accion === "EXTRAER_XML") {
            const seleccion = context.document.getSelection();
            const xmlResult = seleccion.getOoxml();
            await context.sync();
            
            const body = context.document.body;
            body.insertParagraph("--- COPY START ---", "End");
            body.insertParagraph(xmlResult.value, "End");
            body.insertParagraph("--- COPY END ---", "End");
            await context.sync();
        }

    }).catch(error => {
        console.error("Error crítico en Word.run:", error);
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