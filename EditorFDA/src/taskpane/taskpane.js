/* global document, Office, Word */

// 1. CONFIGURACIÓN DE OPCIONES (Global)
const OPCIONES_REVISION = {
    "Interna": ["A", "B"],
    "Codelco": ["B", "C", "D"],
    "Fase":    ["P", "Q", "R"]
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Office.js listo. Iniciando Taskpane...");

        // A. ASIGNAR EVENTOS (Botones y Selects)
        asignarEventos();

        // B. INICIALIZAR INTERFAZ (Fecha y Dropdowns)
        // Esto debe correr SIEMPRE, haya o no datos en memoria
        establecerFechaHoy();
        actualizarListaRevisiones();

        // C. CARGAR DATOS DEL PROYECTO (Con protección anti-errores)
        cargarDatosDeMemoria();
    }
});

function asignarEventos() {
    // Botón Título
    const btnTitulo = document.getElementById("btnActualizarTitulo");
    if (btnTitulo) btnTitulo.onclick = actualizarTituloDocumento;

    // Botón Revisión
    const btnRev = document.getElementById("btnActualizarRevision");
    if (btnRev) btnRev.onclick = actualizarDatosRevision;

    // Dropdown "Emitido Para" (Evento de cambio)
    const ddlEmitido = document.getElementById("ddlEmitidoPara");
    if (ddlEmitido) {
        ddlEmitido.onchange = function() {
            console.log("Cambio detectado en Emitido Para");
            actualizarListaRevisiones();
        };
    }
}

// ---------------------------------------------
// LÓGICA DE INTERFAZ (SECCIÓN 3)
// ---------------------------------------------

function establecerFechaHoy() {
    try {
        const txtFecha = document.getElementById("txtFecha");
        if (txtFecha) {
            const hoy = new Date();
            const dia = String(hoy.getDate()).padStart(2, '0');
            const mes = String(hoy.getMonth() + 1).padStart(2, '0');
            const anio = hoy.getFullYear();
            
            txtFecha.value = `${dia}-${mes}-${anio}`;
            console.log("Fecha establecida:", txtFecha.value);
        }
    } catch (e) {
        console.error("Error al poner la fecha:", e);
    }
}

function actualizarListaRevisiones() {
    try {
        const ddlEmitido = document.getElementById("ddlEmitidoPara");
        const ddlRevision = document.getElementById("ddlRevision");
        
        if (!ddlEmitido || !ddlRevision) return;

        // 1. Ver qué eligió el usuario
        const seleccion = ddlEmitido.value; 
        console.log("Opción seleccionada:", seleccion);
        
        // 2. Obtener las opciones correspondientes
        const opciones = OPCIONES_REVISION[seleccion] || [];

        // 3. Limpiar y llenar el segundo dropdown
        ddlRevision.innerHTML = "";
        
        if (opciones.length === 0) {
            // Opción por defecto si no hay datos
            const opt = document.createElement("option");
            opt.text = "--";
            ddlRevision.appendChild(opt);
        } else {
            opciones.forEach(letra => {
                const opt = document.createElement("option");
                opt.value = letra;
                opt.text = letra;
                ddlRevision.appendChild(opt);
            });
        }
    } catch (e) {
        console.error("Error actualizando lista de revisiones:", e);
    }
}

async function actualizarDatosRevision() {
    const msgLabel = document.getElementById("mensajeEstado");
    
    // Obtener valores
    const ddlEmitido = document.getElementById("ddlEmitidoPara");
    const textoEmitido = ddlEmitido.options[ddlEmitido.selectedIndex].text;
    
    const textoRevision = document.getElementById("ddlRevision").value;
    const textoFecha = document.getElementById("txtFecha").value;

    if(msgLabel) msgLabel.textContent = "Actualizando datos de revisión...";

    await Word.run(async (context) => {
        // Mapeo de TAGs vs VALORES
        const itemsAActualizar = [
            { tag: "ccEmision",  valor: textoEmitido },
            { tag: "ccRevision", valor: textoRevision },
            { tag: "ccFecha",    valor: textoFecha }
        ];

        let cambios = 0;
        
        // Iteramos para buscar y reemplazar
        for (let item of itemsAActualizar) {
            // getByTag devuelve una colección (pueden haber varios con el mismo tag)
            const controls = context.document.contentControls.getByTag(item.tag);
            controls.load("items");
            await context.sync();

            if (controls.items.length > 0) {
                controls.items.forEach(cc => {
                    cc.insertText(item.valor, "Replace");
                    cambios++;
                });
            }
        }

        await context.sync();
        
        if (msgLabel) {
            msgLabel.textContent = cambios > 0 
                ? "✅ Revisión actualizada correctamente." 
                : "⚠️ No se encontraron los controles (tags) en el documento.";
        }
    }).catch(error => {
        console.error("Error Word.run:", error);
        if(msgLabel) msgLabel.textContent = "❌ Error al escribir en Word.";
    });
}


// ---------------------------------------------
// LÓGICA DE DATOS Y MEMORIA (SECCIÓN 1 y 2)
// ---------------------------------------------

async function cargarDatosDeMemoria() {
    try {
        console.log("Intentando leer memoria local...");
        const jsonDatos = localStorage.getItem("FDA_ProyectoActual");
        
        if (jsonDatos) {
            const datos = JSON.parse(jsonDatos);
            console.log("Datos cargados:", datos);

            setText("lblCliente",   datos.cliente);
            setText("lblDivision",  datos.division);
            setText("lblContrato",  datos.contrato);
            setText("lblApi",       datos.api);
            setText("lblProyecto",  datos.nombre);
            setText("lblServicio",  datos.servicio || datos.TipoServicio); 

            const inputTitulo = document.getElementById("txtTituloDoc");
            if (inputTitulo) inputTitulo.value = datos.nombre_doc || ""; 
            
            // Intentar escribir en Word (silencioso si falla)
            escribirDatosBaseEnWord(datos).catch(e => console.warn("No se pudo escribir en Word al inicio:", e));

        } else {
            console.log("No se encontraron datos en memoria.");
        }
    } catch (e) {
        // IMPORTANTE: Si falla el localStorage (por el bloqueo), lo capturamos aquí
        // para que NO rompa el resto de la aplicación.
        console.error("Error crítico leyendo memoria (posible bloqueo de navegador):", e);
        const msg = document.getElementById("mensajeEstado");
        if(msg) msg.textContent = "⚠️ Error de acceso a datos locales. Verifique configuración de cookies.";
    }
}

async function escribirDatosBaseEnWord(datos) {
    await Word.run(async (context) => {
        const tagsMapa = [
            { t: "ccCliente",              v: datos.cliente }, 
            { t: "ccDivisión",             v: datos.division },
            { t: "ccServicios",            v: datos.servicio },
            { t: "ccContrato",             v: datos.contrato },
            { t: "ccAPI",                  v: datos.api },
            { t: "ccProyecto",             v: datos.nombre },
            { t: "ccCliente_encabezado",   v: datos.cliente },
            { t: "ccNProyecto_Encabezado", v: datos.nombre }
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
    const txtTitulo = document.getElementById("txtTituloDoc");
    const nuevoTitulo = txtTitulo ? txtTitulo.value : "";
    const msgLabel = document.getElementById("mensajeEstado");
    
    if(!nuevoTitulo) return;
    if(msgLabel) msgLabel.textContent = "Actualizando título...";

    await Word.run(async (context) => {
        const controls = context.document.contentControls.getByTag("ccNombreDoc");
        controls.load("items");
        await context.sync();
        if (controls.items.length > 0) {
             controls.items.forEach(cc => cc.insertText(nuevoTitulo, "Replace"));
             if(msgLabel) msgLabel.textContent = "✅ Título actualizado.";
        } else {
             if(msgLabel) msgLabel.textContent = "⚠️ Tag 'ccNombreDoc' no encontrado.";
        }
    });
}

function setText(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val || "--";
}