/* global document, Office, Word, fetch, localStorage */

// 1. CONFIGURACIÓN
const URL_POWER_AUTOMATE = "https://defaultef8b3c00d87343e58b66d56c25f2bd.fe.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d88cc5b40d1b48bfa41f130960371fe1/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QAwT8H-2RLeYuIvy4ISgzt0sXfcBX0JGvjjR_3l1V_Y"; 

const OPCIONES_REVISION = {
    "Interna": ["A", "B"],
    "Codelco": ["B", "C", "D"],
    "Fase":    ["P", "Q", "R"]
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Office.js listo. Iniciando Taskpane...");

        asignarEventos();
        establecerFechaHoy();
        actualizarListaRevisiones();
        
        // CORRECCIÓN: Volvemos a activar la carga, pero ahora la función
        // será inteligente y no cargará basura si el documento está en blanco.
        cargarDatosDeMemoria(); 
    }
});

function asignarEventos() {
    const ddlDocs = document.getElementById("ddlDocumentos");
    if (ddlDocs) ddlDocs.onchange = insertarDocumentoSeleccionado;

    const btnRev = document.getElementById("btnActualizarRevision");
    if (btnRev) btnRev.onclick = actualizarDatosRevision;

    const ddlEmitido = document.getElementById("ddlEmitidoPara");
    if (ddlEmitido) ddlEmitido.onchange = actualizarListaRevisiones;
}

// ... (Tu función cargarDocumentosDesdeAzure se mantiene IGUAL) ...
// ... (Tu función insertarDocumentoSeleccionado se mantiene IGUAL) ...

// ---------------------------------------------
// 3. LÓGICA DE DATOS Y MEMORIA (CORREGIDA)
// ---------------------------------------------

async function cargarDatosDeMemoria() {
    try {
        // PASO 1: Verificar si es un documento válido antes de cargar nada
        // Evita que un Word en blanco muestre datos viejos
        let esDocumentoValido = false;

        await Word.run(async (context) => {
            // Buscamos un control clave, ej: ccCliente
            const ccCheck = context.document.contentControls.getByTag("ccCliente");
            ccCheck.load("items");
            await context.sync();

            // Si hay items, es una plantilla nuestra. Si es 0, es un Word en blanco.
            if (ccCheck.items.length > 0) {
                esDocumentoValido = true;
            }
        });

        if (!esDocumentoValido) {
            console.log("Documento en blanco o sin tags. No se cargan datos de memoria.");
            return; // SALIMOS: El lateral quedará con "--"
        }

        // PASO 2: Si es válido, cargamos los datos
        const jsonDatos = localStorage.getItem("FDA_ProyectoActual");
        
        if (jsonDatos) {
            const datos = JSON.parse(jsonDatos);
            
            // Llenar el Lateral (HTML) - SIN API NI SERVICIO
            setText("lblCliente",   datos.cliente);
            setText("lblDivision",  datos.division);
            setText("lblContrato",  datos.contrato);
            // ELIMINADOS VISUALMENTE: API y Servicio
            setText("lblProyecto",  datos.nombre);

            // Escribir en el Word
            escribirDatosBaseEnWord(datos).catch(e => console.warn(e));

            // Cargar lista de documentos desde Azure
            const idProyecto = datos.id; 
            if (idProyecto) {
                cargarDocumentosDesdeAzure(idProyecto);
            }
        }
    } catch (e) {
        console.error("Error leyendo memoria o verificando documento:", e);
    }
}

// ---------------------------------------------
// 4. FUNCIONES AUXILIARES (MODIFICADA)
// ---------------------------------------------

async function escribirDatosBaseEnWord(datos) {
    await Word.run(async (context) => {
        const tagsMapa = [
            { t: "ccCliente",              v: datos.cliente },
            { t: "ccDivisión",             v: datos.division },
            // ELIMINADO: ccServicios
            { t: "ccContrato",             v: datos.contrato },
            // ELIMINADO: ccAPI
            { t: "ccProyecto",             v: datos.nombre },
            { t: "ccCliente_encabezado",   v: datos.cliente },
            { t: "ccNProyecto_Encabezado", v: datos.nombre }
        ];

        // Lógica optimizada de escritura
        let controlesCargados = [];
        for (let item of tagsMapa) {
            if (!item.v) continue;
            let ccs = context.document.contentControls.getByTag(item.t);
            ccs.load("items");
            controlesCargados.push({ ccs: ccs, valor: item.v });
        }

        await context.sync();

        for (let obj of controlesCargados) {
            if (obj.ccs.items.length > 0) {
                obj.ccs.items.forEach(cc => cc.insertText(obj.valor, "Replace"));
            }
        }
    });
}

// ... (Resto de funciones: setText, establecerFechaHoy, actualizarListaRevisiones... IGUALES) ...