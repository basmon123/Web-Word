/* global Office, document, fetch */

// --- CONFIGURACIÓN: Tu base de datos de tablas en GitHub ---
const URL_TABLAS_JSON = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/tablas.json";
let tablasCache = []; // Aquí guardamos la lista para no descargarla a cada rato

Office.onReady(async () => {
    
    // ==========================================
    // 1. INICIALIZAR LA BIBLIOTECA (PLANTILLAS)
    // ==========================================
    await cargarTablasDesdeNube();

    const btnPlantilla = document.getElementById("btnInsertarPlantilla");
    const ddlPlantillas = document.getElementById("ddlPlantillasTablas");

    // Botón para insertar la plantilla seleccionada
    if (btnPlantilla) {
        btnPlantilla.onclick = insertarTablaSeleccionada;
    }

    // Evento para mostrar la descripción cuando cambias de tabla en la lista
    if (ddlPlantillas) {
        ddlPlantillas.onchange = function() {
            const codigo = this.value;
            // Buscamos en la memoria los datos de esa tabla
            const tabla = tablasCache.find(t => t.codigo_xml === codigo);
            const info = document.getElementById("infoTabla");
            
            if (info) {
                // Si encontramos la tabla, mostramos su descripción
                info.textContent = tabla ? (tabla.descripcion || "") : "";
            }
        };
    }

    // ==========================================
    // 2. LÓGICA DE TABLA SIMPLE (FILAS/COLS)
    // ==========================================
    const inputFilas = document.getElementById("txtFilas");
    const inputCols = document.getElementById("txtCols");
    const btnSimple = document.getElementById("btnInsertar"); // Botón "CREAR TABLA"

    if (inputFilas) inputFilas.oninput = actualizarPreview;
    if (inputCols) inputCols.oninput = actualizarPreview;
    
    if (btnSimple) {
        btnSimple.onclick = enviarDatosSimpleAWord;
    }

    actualizarPreview(); // Dibujar la vista previa inicial

    // ==========================================
    // 3. HERRAMIENTA DESARROLLADOR (ESCANER)
    // ==========================================
    const btnExtraer = document.getElementById("btnExtraerCodigo");
    if (btnExtraer) {
        btnExtraer.onclick = function() {
            // Enviamos la orden de extracción a commands.js
            const config = { accion: "EXTRAER_XML" };
            Office.context.ui.messageParent(JSON.stringify(config));
        };
    }
});

// ---------------------------------------------------------
// FUNCIONES DE LA BIBLIOTECA (CARGA DE JSON)
// ---------------------------------------------------------

async function cargarTablasDesdeNube() {
    const ddl = document.getElementById("ddlPlantillasTablas");
    if (!ddl) return; // Si no existe el dropdown, salimos (por si acaso)

    try {
        // Usamos un timestamp (?t=...) para evitar que el navegador guarde versiones viejas
        const response = await fetch(URL_TABLAS_JSON + "?t=" + new Date().getTime());
        
        if (!response.ok) throw new Error("No se pudo conectar con GitHub");

        tablasCache = await response.json();

        // Limpiamos y llenamos el Select
        ddl.innerHTML = '<option value="">-- Seleccione una Tabla --</option>';
        
        tablasCache.forEach(t => {
            let opt = document.createElement("option");
            opt.text = t.nombre;       // Lo que ve el usuario
            opt.value = t.codigo_xml;  // El código secreto (ADN)
            ddl.appendChild(opt);
        });

    } catch (e) {
        console.error(e);
        ddl.innerHTML = '<option>Error cargando lista (Ver Consola)</option>';
    }
}

function insertarTablaSeleccionada() {
    const ddl = document.getElementById("ddlPlantillasTablas");
    const xmlCode = ddl.value;
    
    // Si no seleccionó nada (valor vacío), no hacemos nada
    if(!xmlCode) return;

    // Enviamos el mensaje a commands.js
    const config = {
        accion: "INSERTAR_XML",
        xml: xmlCode
    };
    Office.context.ui.messageParent(JSON.stringify(config));
}

// ---------------------------------------------------------
// FUNCIONES DE TABLA SIMPLE
// ---------------------------------------------------------

function actualizarPreview() {
    const f = document.getElementById("txtFilas").value;
    const c = document.getElementById("txtCols").value;
    const tabla = document.getElementById("tablaPreview");
    
    if (!tabla) return;

    tabla.innerHTML = "";
    for(let i=0; i<f; i++){
        let row = tabla.insertRow();
        for(let j=0; j<c; j++){
            let cell = row.insertCell();
            cell.innerHTML = "&nbsp;"; 
        }
    }
}

function enviarDatosSimpleAWord() {
    // Recopilar datos de los inputs
    const filas = document.getElementById("txtFilas").value;
    const columnas = document.getElementById("txtCols").value;

    const config = {
        accion: "INSERTAR", // Importante: Le dice a commands.js que use lógica simple
        filas: filas,
        columnas: columnas
    };

    // Enviar mensaje a la ventana padre
    Office.context.ui.messageParent(JSON.stringify(config));
}