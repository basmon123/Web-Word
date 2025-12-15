/* global Office, document, fetch */

const URL_TABLAS_JSON = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/tablas.json";
let tablasCache = [];

Office.onReady(async () => {
    // 1. Cargar Biblioteca
    await cargarTablasDesdeNube();

    // 2. Configurar Inputs para Vista Previa (RESTAURADO)
    const inputFilas = document.getElementById("txtFilas");
    const inputCols = document.getElementById("txtCols");
    if (inputFilas) inputFilas.oninput = actualizarPreview;
    if (inputCols) inputCols.oninput = actualizarPreview;

    // 3. Configurar Botones
    const btnSimple = document.getElementById("btnInsertar");
    if (btnSimple) btnSimple.onclick = enviarDatosSimple;

    const btnPlantilla = document.getElementById("btnInsertarPlantilla");
    if (btnPlantilla) btnPlantilla.onclick = enviarDatosPlantilla;
    
    const btnScan = document.getElementById("btnExtraerCodigo");
    if(btnScan) {
        btnScan.onclick = function() {
            // Enviamos la orden de escanear sin cerrar ventana
            Office.context.ui.messageParent(JSON.stringify({ accion: "EXTRAER_XML" }));
        };
    }

    // 4. Dibujar vista previa inicial (RESTAURADO)
    actualizarPreview();
});

// --- FUNCIÓN: Actualizar Vista Previa Visual (RESTAURADA) ---
function actualizarPreview() {
    const fInput = document.getElementById("txtFilas");
    const cInput = document.getElementById("txtCols");
    const tabla = document.getElementById("tablaPreview");
    
    if (!tabla || !fInput || !cInput) return;

    const f = parseInt(fInput.value) || 1;
    const c = parseInt(cInput.value) || 1;

    tabla.innerHTML = "";
    for(let i=0; i<f; i++){
        let row = tabla.insertRow();
        for(let j=0; j<c; j++){
            let cell = row.insertCell();
            // Usamos un caracter visible pequeño para que se note la celda
            cell.innerHTML = "·"; 
        }
    }
}

// --- FUNCIONES DE ENVÍO ---
function enviarDatosSimple() {
    const filas = document.getElementById("txtFilas").value || 3;
    const cols = document.getElementById("txtCols").value || 3;
    const config = { accion: "INSERTAR", filas: filas, columnas: cols };
    Office.context.ui.messageParent(JSON.stringify(config));
}

function enviarDatosPlantilla() {
    const xml = document.getElementById("ddlPlantillasTablas").value;
    if(!xml) return;
    const config = { accion: "INSERTAR_XML", xml: xml };
    Office.context.ui.messageParent(JSON.stringify(config));
}

// --- FUNCIÓN: Cargar JSON ---
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
    }
}