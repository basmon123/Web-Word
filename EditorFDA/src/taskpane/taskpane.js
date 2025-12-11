/* global document, Office, Word, fetch, localStorage */

// 1. CONFIGURACIÓN (Global)
// -----------------------------------------------------------------------------
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
    // --- Evento del Dropdown ---
    const ddlDocs = document.getElementById("ddlDocumentos");
    if (ddlDocs) {
        ddlDocs.onchange = insertarDocumentoSeleccionado;
    }

    // Botón Revisión
    const btnRev = document.getElementById("btnActualizarRevision");
    if (btnRev) btnRev.onclick = actualizarDatosRevision;

    // Dropdown "Emitido Para"
    const ddlEmitido = document.getElementById("ddlEmitidoPara");
    if (ddlEmitido) {
        ddlEmitido.onchange = actualizarListaRevisiones;
    }
}

// ---------------------------------------------
// 2. LÓGICA DE AZURE (Con Filtro Anti-Duplicados)
// ---------------------------------------------

async function cargarDocumentosDesdeAzure(idProyecto) {
    const ddl = document.getElementById("ddlDocumentos");
    if (!ddl) return;

    ddl.innerHTML = "<option>Cargando lista de documentos...</option>";

    try {
        console.log(`Consultando documentos para proyecto ID: ${idProyecto}`);

        const response = await fetch(URL_POWER_AUTOMATE, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ "codigoProyecto": idProyecto }) 
        });

        if (!response.ok) throw new Error("Error de conexión con Power Automate");

        const listaCruda = await response.json();
        console.log(`Recibidos ${listaCruda.length} registros. Filtrando duplicados...`);

        // --- FILTRO DE DUPLICADOS ---
        const documentosUnicos = [];
        const codigosVistos = new Set();

        listaCruda.forEach(doc => {
            // Usamos 'codFDA' como identificador único. Si está vacío, usamos el 'Nombre'.
            // Esto asegura que "Cierre Operacional" aparezca solo una vez.
            const idUnico = doc.codFDA || doc.Nombre; 

            if (!codigosVistos.has(idUnico)) {
                codigosVistos.add(idUnico);
                documentosUnicos.push(doc);
            }
        });

        console.log(`Lista final limpia: ${documentosUnicos.length} documentos.`);

        // Limpiar dropdown
        ddl.innerHTML = "";

        if (documentosUnicos.length === 0) {
            ddl.innerHTML = "<option>No se encontraron documentos</option>";
            return;
        }

        // Opción por defecto
        const optDef = document.createElement("option");
        optDef.text = "-- Seleccione un documento --";
        optDef.value = "";
        ddl.appendChild(optDef);

        // Llenar el dropdown
        documentosUnicos.forEach(doc => {
            const opt = document.createElement("option");
            
            // OJO: Usamos las mayúsculas exactas de tu JSON
            opt.text = doc.Nombre; 
            
            // Guardamos el código cliente en el valor (si es null ponemos guion)
            opt.value = doc.codCliente || "SIN-CODIGO"; 
            
            ddl.appendChild(opt);
        });

    } catch (error) {
        console.error("Error cargando documentos:", error);
        ddl.innerHTML = "<option>Error al cargar lista</option>";
    }
}

async function insertarDocumentoSeleccionado() {
    const ddl = document.getElementById("ddlDocumentos");
    
    const nombreDoc = ddl.options[ddl.selectedIndex].text;
    let codigoCliente = ddl.value;

    if (!codigoCliente || codigoCliente === "") return;
    if (codigoCliente === "SIN-CODIGO") codigoCliente = "N/A"; // Estético para Word

    console.log(`Insertando -> Nombre: ${nombreDoc}, Código: ${codigoCliente}`);

    await Word.run(async (context) => {
        // 1. Insertar Nombre
        const ctrlsNombre = context.document.contentControls.getByTag("ccNombreDoc");
        ctrlsNombre.load("items");

        // 2. Insertar Código Cliente
        const ctrlsCodigo = context.document.contentControls.getByTag("ccCodigoCliente");
        ctrlsCodigo.load("items");

        await context.sync();

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
// 3. LÓGICA DE DATOS Y MEMORIA
// ---------------------------------------------

async function cargarDatosDeMemoria() {
    try {
        const jsonDatos = localStorage.getItem("FDA_ProyectoActual");
        
        if (jsonDatos) {
            const datos = JSON.parse(jsonDatos);
            
            // UI Labels
            setText("lblCliente",   datos.cliente);
            setText("lblDivision",  datos.division);
            setText("lblContrato",  datos.contrato);
            setText("lblApi",       datos.api);
            setText("lblProyecto",  datos.nombre);
            setText("lblServicio",  datos.servicio || datos.TipoServicio);

            // Escribir en Word (Datos base)
            escribirDatosBaseEnWord(datos).catch(e => console.warn(e));

            // --- CORRECCIÓN CLAVE ---
            // Usamos 'id' porque ya confirmamos que ese es el dato correcto (7827)
            const idProyecto = datos.id; 
            
            if (idProyecto) {
                cargarDocumentosDesdeAzure(idProyecto);
            } else {
                console.warn("El proyecto en memoria no tiene un ID válido.");
                const ddl = document.getElementById("ddlDocumentos");
                if(ddl) ddl.innerHTML = "<option>Seleccione un proyecto primero</option>";
            }

        }
    } catch (e) {
        console.error("Error leyendo memoria:", e);
    }
}

// ---------------------------------------------
// 4. FUNCIONES AUXILIARES (Iguales)
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