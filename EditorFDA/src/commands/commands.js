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

async function procesarMensajeTabla(arg) {
    let datos;
    try { datos = JSON.parse(arg.message); } catch (e) { return; }
    
    if (datos.accion !== "EXTRAER_XML") {
        if (dialogGenerador) dialogGenerador.close();
    }

    await Word.run(async (context) => {
        
        // ==========================================
        // CASO A: INSERTAR MANUAL (A PRUEBA DE BALAS)
        // ==========================================
        if (datos.accion === "INSERTAR") {
            const seleccion = context.document.getSelection();
            
            let f = parseInt(datos.filas) || 3;
            let c = parseInt(datos.columnas) || 3;
            
            let matriz = [];
            for(let i=0; i<f; i++) matriz.push(new Array(c).fill(" "));

            try {
                // 1. CREAR TABLA
                const tabla = seleccion.insertTable(f, c, "After", matriz);
                await context.sync(); // Pausa obligatoria

                // 2. FORMATO GENERAL (Fuente)
                // Usamos .set() que es más limpio y seguro
                tabla.getRange().font.set({
                    name: "Arial",
                    size: 12,
                    color: "black"
                });

                // 3. ENCABEZADO (MÉTODO UNIVERSAL)
                // getFirst() existe en todas las versiones de Word API
                const primeraFila = tabla.rows.getFirst();
                
                primeraFila.shading.color = "#1F4E78"; // Azul Oscuro
                primeraFila.font.set({
                    color: "white",
                    bold: true
                });

                // 4. BORDES (CON SEGURIDAD)
                // Envolvemos esto en try/catch por si tu Word es muy antiguo para 'getBorder'
                try {
                    tabla.getBorder("Top").type = "Single";
                    tabla.getBorder("Bottom").type = "Single";
                    tabla.getBorder("Left").type = "Single";
                    tabla.getBorder("Right").type = "Single";
                    tabla.getBorder("InsideHorizontal").type = "Single";
                    tabla.getBorder("InsideVertical").type = "Single";
                } catch (eBordes) {
                    // Si fallan los bordes manuales, intentamos aplicar un estilo nativo como Plan B
                    console.log("No se pudo usar getBorder, intentando estilo...");
                    try { tabla.style = "Table Grid"; } catch(e){}
                }

                // 5. FINALIZAR
                tabla.autofitWindow();
                await context.sync();
                
            } catch (error) {
                const body = context.document.body;
                body.insertParagraph("❌ ERROR CRÍTICO: " + error.message, "Start");
                await context.sync();
            }
        } 
        
        // ==========================================
        // CASO B: INSERTAR PLANTILLA (XML)
        // ==========================================
        else if (datos.accion === "INSERTAR_XML") {
            const seleccion = context.document.getSelection();
            try {
                seleccion.insertOoxml(datos.xml, "After");
                seleccion.insertParagraph("", "After"); 
                await context.sync();
            } catch (errorXML) {
                const body = context.document.body;
                body.insertParagraph("❌ ERROR XML: " + errorXML.message, "Start");
                await context.sync();
            }
        }

        // ==========================================
        // CASO C: ESCANEAR
        // ==========================================
        else if (datos.accion === "EXTRAER_XML") {
            const seleccion = context.document.getSelection();
            const xmlResult = seleccion.getOoxml();
            await context.sync();
            
            const body = context.document.body;
            body.insertParagraph("--- COPY TABLE START ---", "End");
            body.insertParagraph(xmlResult.value, "End");
            body.insertParagraph("--- COPY TABLE END ---", "End");
            await context.sync();
        }

    }).catch(error => {
        console.error("Error crítico:", error);
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
        // CASO 1: INSERTAR PLANTILLA (LAVADO DE CÓDIGO)
        // ==========================================
      if (datos.accion === "INSERTAR_XML") {
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
        // CASO 2: ESCANEAR (DEVELOPER)
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


