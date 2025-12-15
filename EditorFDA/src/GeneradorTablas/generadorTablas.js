/* global Office, document, fetch */

const URL_TABLAS_JSON = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/tablas.json";
let tablasCache = [];

Office.onReady(async () => {
    // 1. Cargar Biblioteca
    await cargarTablasDesdeNube();

    // 2. Botones de Insertar
    const btnSimple = document.getElementById("btnInsertar");
    if (btnSimple) btnSimple.onclick = enviarDatosSimple;

    const btnPlantilla = document.getElementById("btnInsertarPlantilla");
    if (btnPlantilla) btnPlantilla.onclick = enviarDatosPlantilla;
    
    // 3. BOTÓN ESCÁNER (MODIFICADO PARA DAR FEEDBACK VISUAL)
    const btnScan = document.getElementById("btnExtraerCodigo");
    if(btnScan) {
        btnScan.onclick = function() {
            // Cambiamos el texto del botón para que sepas que hizo clic
            btnScan.innerText = "⏳ Enviando orden...";
            btnScan.style.backgroundColor = "#ccc";
            
            // Enviamos la orden
            Office.context.ui.messageParent(JSON.stringify({ accion: "EXTRAER_XML" }));

            // Restauramos el botón a los 2 segundos
            setTimeout(() => {
                btnScan.innerText = "ESCANEAR TABLA";
                btnScan.style.backgroundColor = "orange";
            }, 2000);
        };
    }
});

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

async function cargarTablasDesdeNube() {
    const ddl = document.getElementById("ddlPlantillasTablas");
    if (!ddl) return;
    try {
        const response = await fetch(URL_TABLAS_JSON + "?t=" + new Date().getTime());
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