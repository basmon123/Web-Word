/* global Office, document, fetch */

const URL_GRAFICOS_JSON = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/graficos.json";
let graficosCache = [];

Office.onReady(async () => {
    await cargarBiblioteca();

    // Solo configuramos el botón de Plantilla
    const btnPlantilla = document.getElementById("btnInsertarPlantilla");
    if (btnPlantilla) btnPlantilla.onclick = insertarPlantilla;

    // Y el botón de Escáner
    const btnScan = document.getElementById("btnExtraerCodigo");
    if(btnScan) {
        btnScan.onclick = function() {
            const originalText = btnScan.innerText;
            btnScan.innerText = "⏳ Escaneando...";
            Office.context.ui.messageParent(JSON.stringify({ accion: "EXTRAER_XML" }));
            setTimeout(() => { btnScan.innerText = originalText; }, 2000);
        };
    }
});

function insertarPlantilla() {
    const xml = document.getElementById("ddlPlantillasGraficos").value;
    if(!xml) return;
    // Enviamos el XML al sistema
    const config = { accion: "INSERTAR_XML", xml: xml };
    Office.context.ui.messageParent(JSON.stringify(config));
}

async function cargarBiblioteca() {
    const ddl = document.getElementById("ddlPlantillasGraficos");
    try {
        const response = await fetch(URL_GRAFICOS_JSON + "?t=" + new Date().getTime());
        if (!response.ok) throw new Error("Error conexión");
        graficosCache = await response.json();
        
        ddl.innerHTML = '<option value="">-- Seleccione Gráfico --</option>';
        graficosCache.forEach(g => {
            let opt = document.createElement("option");
            opt.text = g.nombre;
            opt.value = g.codigo_xml;
            ddl.appendChild(opt);
        });
    } catch (e) {
        ddl.innerHTML = '<option>Error cargando lista</option>';
    }
}