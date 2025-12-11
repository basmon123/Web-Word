/* global document, Office, Word, fetch, localStorage */

// 1. CONFIGURACIÓN (Global)
// -----------------------------------------------------------------------------
// ¡IMPORTANTE! Pega aquí la URL que copiaste del paso 1 de tu flujo de Power Automate
const URL_POWER_AUTOMATE = "https://defaultef8b3c00d87343e58b66d56c25f2bd.fe.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d88cc5b40d1b48bfa41f130960371fe1/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QAwT8H-2RLeYuIvy4ISgzt0sXfcBX0JGvjjR_3l1V_Y"; 

const OPCIONES_REVISION = {
    "Interna": ["A", "B"],
    "Codelco": ["B", "C", "D"],
    "Fase":    ["P", "Q", "R"]
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Office.js listo. Iniciando Taskpane...");

        // A. ASIGNAR EVENTOS
        asignarEventos();

        // B. INICIALIZAR INTERFAZ
        establecerFechaHoy();
        actualizarListaRevisiones();

        // C. CARGAR DATOS (Memoria + Azure)
        cargarDatosDeMemoria();
    }
});

function asignarEventos() {
    // --- NUEVO: Evento para el desplegable de documentos ---
    const ddlDocs = document.getElementById("ddlDocumentos");
    if (ddlDocs) {
        // Cuando cambias la selección, se inserta en el Word
        ddlDocs.onchange = insertarDocumentoSeleccionado;
    }

    // Botón Revisión (Se mantiene igual)
    const btnRev = document.getElementById("btnActualizarRevision");
    if (btnRev) btnRev.onclick = actualizarDatosRevision;

    // Dropdown "Emitido Para" (Se mantiene igual)
    const ddlEmitido = document.getElementById("ddlEmitidoPara");
    if (ddlEmitido) {
        ddlEmitido.onchange = actualizarListaRevisiones;
    }
}

// ---------------------------------------------
// 2. NUEVA LÓGICA: CONEXIÓN CON AZURE
// ---------------------------------------------

async function cargarDocumentosDesdeAzure(idProyecto) {
    const ddl = document.getElementById("ddlDocumentos");
    if (!ddl) return;

    // Mostrar estado de carga
    ddl.innerHTML = "<option>Cargando lista de documentos...</option>";

    try {
        console.log(`Consultando documentos para proyecto ID: ${idProyecto}`);

        // Llamada a Power Automate (Tu API)
        const response = await fetch(URL_POWER_AUTOMATE, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ "codigoProyecto": idProyecto }) 
        });

        if (!response.ok) throw new Error("Error de conexión con Power Automate");

        const listaDocumentos = await response.json();
        console.log("Documentos recibidos:", listaDocumentos);

        // Limpiar dropdown
        ddl.innerHTML = "";

        if (!listaDocumentos || listaDocumentos.length === 0) {
            ddl.innerHTML = "<option>No se encontraron documentos</option>";
            return;
        }

        // Opción por defecto
        const optDef = document.createElement("option");
        optDef.text = "-- Seleccione un documento --";
        optDef.value = "";
        ddl.appendChild(optDef);

        // Llenar el dropdown
        listaDocumentos.forEach(doc => {
            const opt = document.createElement("option");
            // Texto visible: Nombre del Documento
            opt.text = doc.NombreDocumento; 
            // Valor oculto: Código del Cliente (necesario para insertar después)
            opt.value = doc.CodigoCliente || "SIN-CODIGO"; 
            ddl.appendChild(opt);
        });

    } catch (error) {
        console.error("Error cargando documentos:", error);
        ddl.innerHTML = "<option>Error al cargar lista</option>";
    }
}

async function insertarDocumentoSeleccionado() {
    const ddl = document.getElementById("ddlDocumentos");
    
    // Obtenemos datos de la selección actual
    const nombreDoc = ddl.options[ddl.selectedIndex].text; // Texto visible
    const codigoCliente = ddl.value; // Valor oculto

    // Si seleccionó la opción por defecto, no hacemos nada
    if (!codigoCliente || codigoCliente === "") return;

    console.log(`Insertando -> Nombre: ${nombreDoc}, Código: ${codigoCliente}`);

    await Word.run(async (context) => {
        // 1. Insertar Nombre del Documento (ccNombreDoc)
        const ctrlsNombre = context.document.contentControls.getByTag("ccNombreDoc");
        ctrlsNombre.load("items");

        // 2. Insertar Código del Cliente (ccCodigoCliente)
        const ctrlsCodigo = context.document.contentControls.getByTag("ccCodigoCliente");
        ctrlsCodigo.load("items");

        await context.sync();

        // Escribimos en Word
        if (ctrlsNombre.items.length > 0) {
            ctrlsNombre.items.forEach(cc => cc.insertText(nombreDoc, "Replace"));
        }

        if (ctrlsCodigo.items.length > 0) {
            ctrlsCodigo.items.forEach(cc => cc.insertText(codigoCliente, "Replace"));
        }

        await context.sync();
    }).catch(e => console.error("Error insertando en Word:", e));
}

// ---------------------------------------------
// 3. LÓGICA DE DATOS Y MEMORIA (Modificada)
// ---------------------------------------------

async function cargarDatosDeMemoria() {
    try {
        console.log("Leyendo memoria local...");
        const jsonDatos = localStorage.getItem("FDA_ProyectoActual");
        
        if (jsonDatos) {
            const datos = JSON.parse(jsonDatos);
            
            // Llenar etiquetas del panel
            setText("lblCliente",   datos.cliente);
            setText("lblDivision",  datos.division);
            setText("lblContrato",  datos.contrato);
            setText("lblApi",       datos.api);
            setText("lblProyecto",  datos.nombre);
            setText("lblServicio",  datos.servicio || datos.TipoServicio);

            // Escribir datos base en el documento Word
            escribirDatosBaseEnWord(datos).catch(e => console.warn(e));

            // --- CRUCE CON AZURE ---
            // Usamos el dato 'api' (o 'id') para buscar los documentos específicos
            const idProyecto = datos.id;
            
            if (idProyecto) {
                cargarDocumentosDesdeAzure(idProyecto);
            } else {
                console.warn("El proyecto en memoria no tiene un ID/API válido.");
            }

        } else {
            console.log("No hay datos en memoria (Proyecto no seleccionado).");
        }
    } catch (e) {
        console.error("Error leyendo memoria:", e);
        const msg = document.getElementById("mensajeEstado");
        if(msg) msg.textContent = "⚠️ Error al leer datos locales.";
    }
}

// ---------------------------------------------
// 4. FUNCIONES AUXILIARES (Sin cambios mayores)
// ---------------------------------------------

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

function setText(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val || "--";
}

function establecerFechaHoy() {
    const txtFecha = document.getElementById("txtFecha");
    if (txtFecha) {
        const hoy = new Date();
        // Formato simple DD-MM-AAAA
        const dia = String(hoy.getDate()).padStart(2, '0');
        const mes = String(hoy.getMonth() + 1).padStart(2, '0');
        const anio = hoy.getFullYear();
        txtFecha.value = `${dia}-${mes}-${anio}`;
    }
}

function actualizarListaRevisiones() {
    const ddlEmitido = document.getElementById("ddlEmitidoPara");
    const ddlRevision = document.getElementById("ddlRevision");
    if (!ddlEmitido || !ddlRevision) return;
    
    const seleccion = ddlEmitido.value;
    const opciones = OPCIONES_REVISION[seleccion] || [];
    
    ddlRevision.innerHTML = "";
    
    if (opciones.length === 0) {
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
}

async function actualizarDatosRevision() {
    const msgLabel = document.getElementById("mensajeEstado");
    if(msgLabel) msgLabel.textContent = "Actualizando...";

    await Word.run(async (context) => {
        const ddlEmitido = document.getElementById("ddlEmitidoPara");
        const textoEmitido = ddlEmitido.options[ddlEmitido.selectedIndex].text;
        
        const textoRevision = document.getElementById("ddlRevision").value;
        const textoFecha = document.getElementById("txtFecha").value;

        const items = [
            { t: "ccEmision",  v: textoEmitido },
            { t: "ccRevision", v: textoRevision },
            { t: "ccFecha",    v: textoFecha }
        ];

        let cambios = 0;
        for (let item of items) {
            const controls = context.document.contentControls.getByTag(item.t);
            controls.load("items");
            await context.sync();
            if (controls.items.length > 0) {
                controls.items.forEach(cc => {
                    cc.insertText(item.v, "Replace");
                    cambios++;
                });
            }
        }
        
        if (msgLabel) {
            msgLabel.textContent = cambios > 0 
                ? "✅ Revisión actualizada." 
                : "⚠️ No se encontraron controles en el documento.";
        }
    }).catch(error => {
        console.error("Error:", error);
        if(msgLabel) msgLabel.textContent = "❌ Error al escribir en Word.";
    });
}