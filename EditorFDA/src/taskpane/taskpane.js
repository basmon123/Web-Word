/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // 1. Configurar botón de actualizar título
        const btn = document.getElementById("btnActualizarTitulo");
        if (btn) btn.onclick = actualizarTituloDocumento;

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
            console.log("Datos encontrados (objeto):", datos);

            // A. Llenar la parte visual (Labels)
            // IMPORTANTE: Aquí usamos los nombres en MINÚSCULA tal cual vienen de catalog.js
            setText("lblCliente",   datos.cliente);    // antes datos.Cliente
            setText("lblDivision",  datos.division);   // antes datos.Division
            setText("lblContrato",  datos.contrato);   // antes datos.NumeroContrato
            setText("lblApi",       datos.api);        // antes datos.NumeroAPI
            setText("lblProyecto",  datos.nombre);     // antes datos.NombreProyecto
            
            // Nota: En tu catalog.js no vi que mapearas "tipo de servicio". 
            // Si existe en el JSON original, agrégalo al map del catalog.js. 
            // Por ahora intentamos leerlo si existe, si no, saldrá "--"
            setText("lblServicio",  datos.servicio || datos.TipoServicio); 

            // B. Pre-llenar el input del título (si existiera una propiedad para ello)
            const inputTitulo = document.getElementById("txtTituloDoc");
            if (inputTitulo) {
                // Si tienes un nombre de doc por defecto, úsalo, sino vacío
                inputTitulo.value = datos.nombre_doc || ""; 
            }

            // C. Escribir en el Word (Content Controls)
            await escribirDatosBaseEnWord(datos);

        } else {
            console.log("No hay datos en memoria.");
            const msg = document.getElementById("mensajeEstado");
            if(msg) msg.textContent = "⚠️ No se detectó selección de proyecto previa.";
        }
    } catch (e) {
        console.error("Error leyendo memoria:", e);
    }
}

// --- ESCRITURA EN WORD ---
async function escribirDatosBaseEnWord(datos) {
    await Word.run(async (context) => {
        // Mapeo corregido: Etiquetas del Word (izquierda) vs Variables del Catalog (derecha)
        const tagsMapa = [
            { t: "ccCliente",              v: datos.cliente }, 
            { t: "ccDivisión",             v: datos.division },
            { t: "ccServicios",            v: datos.servicio }, // Ajustar si agregas servicio al catalog
            { t: "ccContrato",             v: datos.contrato },
            { t: "ccAPI",                  v: datos.api },
            { t: "ccProyecto",             v: datos.nombre },
            // Encabezados (si usas duplicados para encabezados)
            { t: "ccCliente_encabezado",   v: datos.cliente },
            { t: "ccNProyecto_Encabezado", v: datos.nombre }
        ];

        for (let item of tagsMapa) {
            if (!item.v) continue; // Si el dato está vacío, saltamos
            
            // Buscamos el control por TAG
            let ccs = context.document.contentControls.getByTag(item.t);
            ccs.load("items");
            await context.sync();
            
            if (ccs.items.length > 0) {
                // Escribimos en todos los controles con ese tag
                ccs.items.forEach(cc => {
                    cc.insertText(item.v, "Replace"); 
                });
            }
        }
    });
}

async function actualizarTituloDocumento() {
    const txtTitulo = document.getElementById("txtTituloDoc");
    const msgLabel = document.getElementById("mensajeEstado");
    
    if(!txtTitulo) return;
    const nuevoTitulo = txtTitulo.value;

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