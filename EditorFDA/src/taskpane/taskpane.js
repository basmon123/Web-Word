/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // 1. Configurar botón de actualizar título
        document.getElementById("btnActualizarTitulo").onclick = actualizarTituloDocumento;

        // 2. RECUPERAR DATOS AUTOMÁTICAMENTE
        cargarDatosDeMemoria();
    }
});

// --- FUNCIÓN PRINCIPAL: LEER MEMORIA ---
async function cargarDatosDeMemoria() {
    try {
        // Buscamos en el buzón "FDA_ProyectoActual"
        const jsonDatos = localStorage.getItem("FDA_ProyectoActual");
        
        if (jsonDatos) {
            const datos = JSON.parse(jsonDatos);
            console.log("Datos encontrados en memoria:", datos);

            // A. Llenar la parte visual (Labels)
            setText("lblCliente", datos.Cliente);
            setText("lblDivision", datos.Division);
            setText("lblServicio", datos.TipoServicio);
            setText("lblContrato", datos.NumeroContrato);
            setText("lblApi", datos.NumeroAPI);
            setText("lblProyecto", datos.NombreProyecto);

            // B. Pre-llenar el input del título
            const inputTitulo = document.getElementById("txtTituloDoc");
            if (inputTitulo) {
                inputTitulo.value = datos.NombreDoc || ""; 
                inputTitulo.placeholder = "Ej: Informe de Avance";
            }

            // C. Escribir en el Word (Content Controls de solo lectura)
            // Esto asegura que si abriste la plantilla, se llenen los datos duros
            await escribirDatosBaseEnWord(datos);

        } else {
            console.log("No hay datos en memoria.");
            document.getElementById("mensajeEstado").textContent = "⚠️ No se detectó selección de proyecto previa.";
        }
    } catch (e) {
        console.error("Error leyendo memoria:", e);
    }
}

// --- ESCRITURA EN WORD ---
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

function setText(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val || "--";
}