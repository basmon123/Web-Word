/* global document, Office, Word, fetch */

// URL DE TU JSON (El origen de los datos)
const URL_JSON = "https://raw.githubusercontent.com/basmon123/templates/main/data.json"; 

// Variable para guardar los datos descargados
let listaProyectosGlobal = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Eventos
        document.getElementById("btnActualizarTitulo").onclick = actualizarTituloDocumento;
        
        // Evento cuando cambias la selección en la lista
        document.getElementById("ddlProyectos").onchange = alCambiarSeleccion;

        // Iniciar carga de datos
        cargarListaDesdeSharePoint();
    }
});

// --- 1. CARGAR DATOS (JSON) Y LLENAR EL SELECT ---
async function cargarListaDesdeSharePoint() {
    const selector = document.getElementById("ddlProyectos");
    try {
        const response = await fetch(URL_JSON);
        if (!response.ok) throw new Error("Error conectando a SharePoint/GitHub");
        
        const data = await response.json();
        // Ajuste: si tu JSON tiene una propiedad "projects", úsala. Si es un array directo, usa data.
        listaProyectosGlobal = data.projects || data; 

        // Limpiar y llenar el select
        selector.innerHTML = '<option value="">-- Selecciona un proyecto --</option>';
        
        listaProyectosGlobal.forEach((proy, index) => {
            const option = document.createElement("option");
            option.value = index; // Usamos el índice para encontrarlo rápido luego
            option.text = proy.NombreProyecto || "Sin Nombre";
            selector.appendChild(option);
        });

    } catch (error) {
        console.error(error);
        selector.innerHTML = '<option value="">Error cargando datos</option>';
        // (Opcional) Cargar datos locales si falla la red, como vimos antes
    }
}

// --- 2. REACCIONAR A LA SELECCIÓN ---
async function alCambiarSeleccion() {
    const index = document.getElementById("ddlProyectos").value;
    if (index === "") return; // Si selecciona el placeholder, no hacer nada

    const datos = listaProyectosGlobal[index]; // Recuperamos el objeto completo

    // A. Llenar la parte visual (Tu imagen)
    setText("lblCliente", datos.Cliente);
    setText("lblDivision", datos.Division);
    setText("lblServicio", datos.TipoServicio);
    setText("lblContrato", datos.NumeroContrato);
    setText("lblApi", datos.NumeroAPI);
    setText("lblProyecto", datos.NombreProyecto);
    
    // Pre-llenar input editable si existe el dato
    const inputTitulo = document.getElementById("txtTituloDoc");
    if(inputTitulo) inputTitulo.value = datos.NombreDoc || "";

    // B. Escribir en el Word automáticamente
    await escribirDatosBaseEnWord(datos);
}

// --- 3. FUNCIONES DE ESCRITURA EN WORD ---

async function escribirDatosBaseEnWord(datos) {
    await Word.run(async (context) => {
        const tagsMapa = [
            { t: "ccCliente",              v: datos.Cliente }, 
            { t: "ccDivisión",             v: datos.Division },
            { t: "ccServicios",            v: datos.TipoServicio },
            { t: "ccContrato",             v: datos.NumeroContrato },
            { t: "ccAPI",                  v: datos.NumeroAPI },
            { t: "ccProyecto",             v: datos.NombreProyecto }
        ];

        for (let item of tagsMapa) {
            if (!item.v) continue;
            let ccs = context.document.contentControls.getByTag(item.t);
            ccs.load("items");
            await context.sync();
            if (ccs.items.length > 0) {
                ccs.items.forEach(cc => cc.insertText(item.v, "Replace"));
            }
        }
    });
}

async function actualizarTituloDocumento() {
    const nuevoTitulo = document.getElementById("txtTituloDoc").value;
    const msgLabel = document.getElementById("mensajeEstado");
    
    if(!nuevoTitulo) return;
    if(msgLabel) msgLabel.textContent = "Actualizando...";

    await Word.run(async (context) => {
        const controls = context.document.contentControls.getByTag("ccNombreDoc");
        controls.load("items");
        await context.sync();
        
        if (controls.items.length > 0) {
             controls.items.forEach(cc => cc.insertText(nuevoTitulo, "Replace"));
             if(msgLabel) msgLabel.textContent = "✅ Título actualizado.";
        } else {
             if(msgLabel) msgLabel.textContent = "⚠️ No se encontró el control ccNombreDoc.";
        }
    });
}

// Utilitario
function setText(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val || "--";
}