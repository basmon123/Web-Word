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

/* Reemplaza TU función procesarMensajeGrafico por esta versión final */

async function procesarMensajeGrafico(arg) {
    let datos;
    try { datos = JSON.parse(arg.message); } catch (e) { return; }

    if (datos.accion !== "EXTRAER_XML") {
        if (dialogGraficos) dialogGraficos.close();
    }

    await Word.run(async (context) => {
        
// ==========================================
        // CASO 1: INSERTAR ESTÁNDAR (TRUCO XML)
        // ==========================================
        if (datos.accion === "INSERTAR_ESTANDAR") {
            const seleccion = context.document.getSelection();
            
            // Como tu Word no tiene "insertChart", usamos el XML de un gráfico base.
            // Esto funcionará porque tus plantillas ya funcionan.
            try {
                // 1. Limpiamos el XML base por seguridad
                let xmlParaInsertar = XML_GRAFICO_BASE;
                xmlParaInsertar = xmlParaInsertar.replace(/\\r\\n/g, ""); 

                // 2. Insertamos usando la función que SÍ funciona en tu Word
                seleccion.insertOoxml(xmlParaInsertar, "After");
                seleccion.insertParagraph("", "After");
                
                await context.sync();
                
            } catch (error) {
                const body = context.document.body;
                body.insertParagraph("❌ ERROR FATAL: Ni siquiera el XML base funcionó.", "Start");
                body.insertParagraph("Detalle: " + error.message, "Start");
            }
        }
        // ==========================================
        // CASO 2: INSERTAR PLANTILLA (LAVADO DE CÓDIGO)
        // ==========================================
        else if (datos.accion === "INSERTAR_XML") {
            const seleccion = context.document.getSelection();
            
            try {
                // --- FASE DE LIMPIEZA PROFUNDA ---
                let xmlLimpio = datos.xml;

                // 1. Eliminar los "\r\n" literales que rompen el XML (EL GRAN CULPABLE)
                // Esto convierte el texto "\r\n" en nada.
                xmlLimpio = xmlLimpio.replace(/\\r\\n/g, "");
                
                // 2. Corregir URLs rotas (http:\/\/ -> http://)
                xmlLimpio = xmlLimpio.replace(/http:\\\/\\\//g, "http://");
                xmlLimpio = xmlLimpio.replace(/http:\\\//g, "http://");

                // 3. Corregir rutas de paquetes (\/_rels -> /_rels)
                xmlLimpio = xmlLimpio.replace(/\\\/_rels/g, "/_rels");
                xmlLimpio = xmlLimpio.replace(/\\\/word/g, "/word");
                
                // ---------------------------------

                // Insertamos el XML ya lavado
                seleccion.insertOoxml(xmlLimpio, "After");
                seleccion.insertParagraph("", "After"); 
                
                await context.sync();

            } catch (error) {
                const body = context.document.body;
                body.insertParagraph("❌ Error XML (Aún limpiando): " + error.message, "Start");
                await context.sync();
            }
        }

        // ==========================================
        // CASO 3: ESCANEAR (DEVELOPER)
        // ==========================================
        else if (datos.accion === "EXTRAER_XML") {
            const seleccion = context.document.getSelection();
            const xml = seleccion.getOoxml();
            await context.sync();
            
            const body = context.document.body;
            body.insertParagraph("--- COPY CHART START ---", "End");
            body.insertParagraph(xml.value, "End");
            body.insertParagraph("--- COPY CHART END ---", "End");
            await context.sync();
        }

    }).catch(error => console.error("Error Crítico:", error));
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

// =========================================================================
// EL SALVAVIDAS: CÓDIGO XML DE UN GRÁFICO DE COLUMNAS BÁSICO
// (Este código simula ser un gráfico creado nativamente)
// =========================================================================
const XML_GRAFICO_BASE = `<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body><w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/charts/chart1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"><pkg:xmlData><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:lang val="es-ES"/><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="123456"/><c:axId val="123457"/></c:barChart><c:catAx><c:axId val="123456"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="123457"/></c:catAx><c:valAx><c:axId val="123457"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="123456"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart><c:externalData r:id="rId1"/></c:chartSpace></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/charts/_rels/chart1.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/embeddings/Microsoft_Excel_Worksheet.xlsx" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"><pkg:binaryData>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</pkg:binaryData></pkg:part></pkg:package>`;