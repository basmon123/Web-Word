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
        
    // ------------------------------------------
        // A. INSERTAR MANUAL (VERSIÓN PASO A PASO BLINDADA)
        // ------------------------------------------
        if (datos.accion === "INSERTAR") {
            const seleccion = context.document.getSelection();
            
            let f = parseInt(datos.filas) || 3;
            let c = parseInt(datos.columnas) || 3;
            let matriz = [];
            for(let i=0; i<f; i++) matriz.push(new Array(c).fill(" "));

            try {
                // PASO 1: Insertar la tabla básica
                const tabla = seleccion.insertTable(f, c, "After", matriz);
                await context.sync(); // ¡Tabla creada!

                // PASO 2: Obtener el Rango y cargar SOLO el objeto 'format'
                // No intentamos cargar bordes todavía.
                const rangoTabla = tabla.getRange();
                rangoTabla.load("format"); 
                
                // También cargamos las filas para el encabezado
                const filas = tabla.rows;
                filas.load("items");
                
                await context.sync(); // ¡Objeto Format cargado!

                // PASO 3: Ahora que 'format' existe seguro, cargamos 'borders' y 'font'
                // Ya no dará undefined porque format ya está en memoria
                rangoTabla.format.load("borders");
                rangoTabla.font.load("name, size, color");
                
                await context.sync(); // ¡Bordes cargados!

                // PASO 4: APLICAR FORMATO
                
                // A) Fuente General
                rangoTabla.font.name = "Arial";
                rangoTabla.font.size = 12;
                rangoTabla.font.color = "black";

                // B) Bordes (Ahora es seguro acceder)
                const bordes = rangoTabla.format.borders;
                
                // Aplicamos uno por uno para asegurar
                bordes.getItem("InsideHorizontal").style = "Single";
                bordes.getItem("InsideVertical").style = "Single";
                bordes.getItem("EdgeBottom").style = "Single";
                bordes.getItem("EdgeLeft").style = "Single";
                bordes.getItem("EdgeRight").style = "Single";
                bordes.getItem("EdgeTop").style = "Single";

                // C) Encabezado (Fila 0)
                if (filas.items.length > 0) {
                    const encabezado = filas.items[0];
                    encabezado.load("shading, font"); // Cargamos propiedades del encabezado
                    await context.sync(); // Pequeña pausa para asegurar

                    encabezado.shading.color = "#1F4E78"; // Azul Oscuro
                    encabezado.font.color = "white";      // Blanco
                    encabezado.font.bold = true;          // Negrita
                }

                tabla.autofitWindow();
                await context.sync();
                
            } catch (error) {
                context.document.body.insertParagraph("❌ Error Formato Detallado: " + error.message, "Start");
                await context.sync();
            }
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
        
      // ------------------------------------------
        /// A. INSERTAR MANUAL (VERSIÓN NATIVA WORD)
        // ------------------------------------------
        if (datos.accion === "INSERTAR") {
            const seleccion = context.document.getSelection();
            
            let f = parseInt(datos.filas) || 3;
            let c = parseInt(datos.columnas) || 3;
            let matriz = [];
            for(let i=0; i<f; i++) matriz.push(new Array(c).fill(" "));

            try {
                // PASO 1: Insertar la tabla básica
                const tabla = seleccion.insertTable(f, c, "After", matriz);
                await context.sync(); // ¡Tabla creada en memoria!

                // PASO 2: Cargar lo necesario para formatear
                // En Word, NO existe 'format'. Cargamos 'font' directo del rango y 'rows' de la tabla.
                const rangoTabla = tabla.getRange();
                rangoTabla.load("font");
                
                const filas = tabla.rows;
                filas.load("items");
                
                await context.sync(); // Traemos los datos

                // PASO 3: APLICAR FORMATO (SINTAXIS CORRECTA WORD)
                
                // A) Fuente General (Arial 12)
                rangoTabla.font.name = "Arial";
                rangoTabla.font.size = 12;
                rangoTabla.font.color = "black";

                // B) Bordes (La forma correcta en Word es usando getBorder en la tabla)
                // Tipos: 'Single', 'None', 'Double', etc.
                tabla.getBorder("Top").type = "Single";
                tabla.getBorder("Bottom").type = "Single";
                tabla.getBorder("Left").type = "Single";
                tabla.getBorder("Right").type = "Single";
                tabla.getBorder("InsideHorizontal").type = "Single";
                tabla.getBorder("InsideVertical").type = "Single";
                
                // Opcional: Asegurar color negro en bordes
                tabla.getBorder("InsideHorizontal").color = "black"; 
                // (Puedes repetir para los demás si es necesario)

                // C) Encabezado (Fila 0) - Estilo del Jefe
                if (filas.items.length > 0) {
                    const encabezado = filas.items[0];
                    // Cargamos shading (fondo) y font (texto) de esa fila específica
                    encabezado.load("shading, font"); 
                    await context.sync();

                    encabezado.shading.color = "#1F4E78"; // Azul Oscuro
                    encabezado.font.color = "white";      // Blanco
                    encabezado.font.bold = true;          // Negrita
                }

                // Ajuste final
                tabla.autofitWindow();
                await context.sync();
                
            } catch (error) {
                context.document.body.insertParagraph("❌ Error Formato Word: " + error.message, "Start");
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

