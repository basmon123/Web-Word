/* global Office, document, fetch */

// URL de tu base de datos de tablas
const URL_TABLAS_JSON = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/tablas.json";
let tablasCache = [];

Office.onReady(async () => {
    // 1. Cargar Biblioteca al iniciar
    await cargarTablasDesdeNube();

    // 2. Configurar Botón: Crear Manual
    const btnSimple = document.getElementById("btnInsertar");
    if (btnSimple) {
        btnSimple.onclick = enviarDatosSimple;
    }

    // 3. Configurar Botón: Insertar Plantilla
    const btnPlantilla = document.getElementById("btnInsertarPlantilla");
    if (btnPlantilla) {
        btnPlantilla.onclick = enviarDatosPlantilla;
    }
    
    // 4. Configurar Botón: Escáner (Developer)
    const btnScan = document.getElementById("btnExtraerCodigo");
    if(btnScan) {
        btnScan.onclick = function() {
            // Enviamos la orden de escanear sin cerrar ventana
            Office.context.ui.messageParent(JSON.stringify({ accion: "EXTRAER_XML" }));
        };
    }
});

// --- FUNCIÓN: Crear Tabla Manual ---
function enviarDatosSimple() {
    const filas = document.getElementById("txtFilas").value || 3;
    const cols = document.getElementById("txtCols").value || 3;
    
    const config = { 
        accion: "INSERTAR", 
        filas: filas, 
        columnas: cols 
    };
    Office.context.ui.messageParent(JSON.stringify(config));
}

// --- FUNCIÓN: Insertar desde Biblioteca ---
function enviarDatosPlantilla() {
    const ddl = document.getElementById("ddlPlantillasTablas");
    const xml = ddl.value;
    
    if(!xml) return; // Si no seleccionó nada, no hace nada

    const config = { 
        accion: "INSERTAR_XML", 
        xml: xml 
    };
    Office.context.ui.messageParent(JSON.stringify(config));
}

// --- FUNCIÓN: Cargar JSON desde GitHub ---
async function cargarTablasDesdeNube() {
    const ddl = document.getElementById("ddlPlantillasTablas");
    if (!ddl) return;

    try {
        // Usamos timestamp para evitar caché antiguo
        const response = await fetch(URL_TABLAS_JSON + "?t=" + new Date().getTime());
        if (!response.ok) throw new Error("Error conexión");

        tablasCache = await response.json();
        
        // Limpiar y llenar
        ddl.innerHTML = '<option value="">-- Seleccione Plantilla --</option>';
        tablasCache.forEach(t => {
            let opt = document.createElement("option");
            opt.text = t.nombre;
            opt.value = t.codigo_xml; // El valor oculto es el código OOXML
            ddl.appendChild(opt);
        });
    } catch (e) {
        console.error(e);
        ddl.innerHTML = '<option>Error cargando lista</option>';
    }
}