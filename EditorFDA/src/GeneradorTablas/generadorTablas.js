/* global Office, document, fetch, window */

const URL_TABLAS_JSON = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/tablas.json";
let tablasCache = [];

Office.onReady(async () => {
    // 1. Cargar Biblioteca
    await cargarTablasDesdeNube();

    // 2. Configurar Botón: Insertar Manual
    const btnSimple = document.getElementById("btnInsertar");
    if (btnSimple) {
        btnSimple.onclick = enviarDatosSimple;
    }

    // 3. Configurar Botón: Insertar Plantilla
    const btnPlantilla = document.getElementById("btnInsertarPlantilla");
    if (btnPlantilla) {
        btnPlantilla.onclick = enviarDatosPlantilla;
    }
    
    // 4. Configurar Botón: Escáner
    const btnScan = document.getElementById("btnExtraerCodigo");
    if(btnScan) {
        btnScan.onclick = function() {
            // Feedback visual
            const originalText = btnScan.innerText;
            btnScan.innerText = "⏳ Procesando...";
            
            Office.context.ui.messageParent(JSON.stringify({ accion: "EXTRAER_XML" }));

            setTimeout(() => { btnScan.innerText = originalText; }, 2000);
        };
    }

    // 5. Dibujar primera vista previa
    actualizarPreview();
});

// --- FUNCIÓN GLOBAL PARA VISTA PREVIA ---
window.actualizarPreview = function() {
    const fInput = document.getElementById("txtFilas");
    const cInput = document.getElementById("txtCols");
    const tabla = document.getElementById("tablaPreview");
    
    if (!tabla || !fInput || !cInput) return;

    // Aseguramos que sean números válidos
    let f = parseInt(fInput.value);
    let c = parseInt(cInput.value);

    // Límites de seguridad visual
    if (isNaN(f) || f < 1) f = 1;
    if (f > 20) f = 20;
    if (isNaN(c) || c < 1) c = 1;
    if (c > 10) c = 10;

    tabla.innerHTML = "";
    for(let i=0; i<f; i++){
        let row = tabla.insertRow();
        for(let j=0; j<c; j++){
            let cell = row.insertCell();
            // Un espacio vacío para que la celda tenga forma
            cell.innerHTML = "&nbsp;"; 
        }
    }
};

// --- ENVÍO DE DATOS MANUALES ---
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

// --- ENVÍO DE PLANTILLA XML ---
function enviarDatosPlantilla() {
    const ddl = document.getElementById("ddlPlantillasTablas");
    const xml = ddl.value;
    
    if(!xml) return; 

    const config = { 
        accion: "INSERTAR_XML", 
        xml: xml 
    };
    Office.context.ui.messageParent(JSON.stringify(config));
}

// --- CARGA DE DATOS ---
async function cargarTablasDesdeNube() {
    const ddl = document.getElementById("ddlPlantillasTablas");
    if (!ddl) return;
    try {
        const response = await fetch(URL_TABLAS_JSON + "?t=" + new Date().getTime());
        if (!response.ok) throw new Error("Error conexión");
        tablasCache = await response.json();
        
        ddl.innerHTML = '<option value="">-- Seleccione Plantilla --</option>';
        tablasCache.forEach(t => {
            let opt = document.createElement("option");
            opt.text = t.nombre;
            opt.value = t.codigo_xml;
            ddl.appendChild(opt);
        });
    } catch (e) {
        ddl.innerHTML = '<option>Error cargando lista</option>';
        console.error(e);
    }
}