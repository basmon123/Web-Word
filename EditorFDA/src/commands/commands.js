/* global Office, Word */

let dialogCatalogo; 
let dialogGenerador;
let dialogGraficos; // Variable para la ventana de gráficos


Office.onReady(() => {
  console.log("Commands.js listo");
});

// ==========================================
// 1. LÓGICA DE TABLAS
// ==========================================

function abrirVentanaTablas(event) {
    const url = "https://basmon123.github.io/Web-Word/EditorFDA/src/generadorTablas/generadorTablas.html"; 
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
        // CASO 1: INSERTAR MANUAL (CORREGIDO)
        // ==========================================
        if (datos.accion === "INSERTAR") {
            const seleccion = context.document.getSelection();

            // 1. Limpieza y conversión de números
            let f = parseInt(datos.filas);
            let c = parseInt(datos.columnas);
            // Si vienen vacíos o inválidos, usamos 3x3 por seguridad
            if (!f || isNaN(f)) f = 3;
            if (!c || isNaN(c)) c = 3;

            // 2. Crear matriz de datos (Array de Arrays de Strings)
            // Es CRÍTICO que sean strings, por eso ponemos " " y no null
            let matriz = [];
            for(let i=0; i<f; i++) {
                let fila = [];
                for(let j=0; j<c; j++) {
                    fila.push(" "); // Celda vacía con un espacio
                }
                matriz.push(fila);
            }

            // 3. INTENTO DE INSERTAR TABLA
            try {
                // Usamos "After" para que no borre lo que tengas seleccionado, sino que la ponga después
                const tabla = seleccion.insertTable(f, c, "After", matriz);
                
                // Opcional: Le damos un estilo básico de Word para que se vean los bordes
                // Si tu Word está en español, "Table Grid" podría fallar, así que lo envolvemos
                try { tabla.style = "Table Grid"; } catch(e){} 

                tabla.autofitWindow();
                
            } catch (errorTabla) {
                // SI FALLA, ESCRIBIMOS EL ERROR EN EL DOCUMENTO
                const body = context.document.body;
                body.insertParagraph("❌ ERROR AL CREAR TABLA: " + errorTabla.message, "Start");
            }
            
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
// SECCIÓN NUEVA: GESTOR DE GRÁFICOS
/// ==========================================
function abrirVentanaGraficos(event) {
    // Asegúrate de la ruta correcta (Mayúsculas/Minúsculas)
    const url = "https://basmon123.github.io/Web-Word/EditorFDA/src/generadorGraficos/generadorGraficos.html"; 
    
    Office.context.ui.displayDialogAsync(url, { height: 50, width: 30, displayInIframe: true }, 
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Fallo diálogo gráficos:", asyncResult.error.message);
            } else {
                dialogGraficos = asyncResult.value;
                dialogGraficos.addEventHandler(Office.EventType.DialogMessageReceived, procesarMensajeGrafico);
            }
        }
    );
    if(event) event.completed();
}

async function procesarMensajeGrafico(arg) {
    let datos;
    try { datos = JSON.parse(arg.message); } catch (e) { return; }

    // Si no es escáner, cerramos ventana
    if (datos.accion !== "EXTRAER_XML") {
        if (dialogGraficos) dialogGraficos.close();
    }

    await Word.run(async (context) => {
        
        // CASO 1: INSERTAR ESTÁNDAR (Nativo de Word)
        if (datos.accion === "INSERTAR_ESTANDAR") {
            const seleccion = context.document.getSelection();
            
            // Convertimos el texto del dropdown al Enum de Word
            // Ej: "Line" -> Word.ChartType.line
            const tipo = datos.tipoGrafico; 
            
            // Insertamos gráfico con datos de ejemplo ("Auto")
            const grafico = seleccion.insertChart(tipo, "Auto", "Auto");
            
            // Ajustes opcionales
            grafico.height = 300; // Tamaño por defecto razonable
            
            await context.sync();
        }

        // CASO 2: INSERTAR PLANTILLA (XML ESCANEADO)
        else if (datos.accion === "INSERTAR_XML") {
            const seleccion = context.document.getSelection();
            try {
                seleccion.insertOoxml(datos.xml, "After");
                seleccion.insertParagraph("", "After");
                await context.sync();
            } catch (error) {
                context.document.body.insertParagraph("❌ Error al insertar gráfico: " + error.message, "Start");
                await context.sync();
            }
        }

        // CASO 3: ESCANEAR (DEVELOPER)
        else if (datos.accion === "EXTRAER_XML") {
            const seleccion = context.document.getSelection();
            // Obtenemos el ADN (OOXML)
            const xml = seleccion.getOoxml();
            await context.sync();
            
            // Escribimos al final para copiar
            const body = context.document.body;
            body.insertParagraph("--- COPY CHART START ---", "End");
            body.insertParagraph(xml.value, "End");
            body.insertParagraph("--- COPY CHART END ---", "End");
            await context.sync();
        }

    }).catch(error => console.error("Error Gráficos:", error));
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

// --- REGISTRO DEL BOTÓN (AGREGA ESTO AL FINAL JUNTO A LOS OTROS) ---
Office.actions.associate("abrirVentanaGraficos", abrirVentanaGraficos);
