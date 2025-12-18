/* global document, Office, Word, fetch, localStorage */

// 1. CONFIGURACIÓN (Global)
// -----------------------------------------------------------------------------
const URL_POWER_AUTOMATE = "https://defaultef8b3c00d87343e58b66d56c25f2bd.fe.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d88cc5b40d1b48bfa41f130960371fe1/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QAwT8H-2RLeYuIvy4ISgzt0sXfcBX0JGvjjR_3l1V_Y"; 

// Variable global para almacenar las revisiones
let revisions = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Office.js listo. Iniciando Taskpane...");

        asignarEventos();
        cargarDatosDeMemoria();
        
        // Inicializar fecha actual en el input de revisión
        establecerFechaHoyInput();
        
        // Inicializar lógica de revisión (sugerir A)
        setNextLogic('INIT');
    }
});

function asignarEventos() {
    // Dropdown Documentos
    const ddlDocs = document.getElementById("ddlDocumentos");
    if (ddlDocs) ddlDocs.onchange = insertarDocumentoSeleccionado;

    // EVENTOS NUEVA GESTIÓN DE REVISIONES
    
    // 1. Botones Lógicos
    document.getElementById("btnIterar").onclick = () => setNextLogic('ITERATE');
    document.getElementById("btnFase").onclick = () => setNextLogic('PHASE');
    
    // 2. Cambio de Estándar (Codelco vs AMSA)
    document.getElementById("ddlEstandar").onchange = () => setNextLogic('UPDATE_TEXT');

    // 3. Agregar a la lista visual (Tabla HTML)
    document.getElementById("btnAgregarAlista").onclick = addRevisionRow;

    // 4. Botón Final (Escribir en Word)
    document.getElementById("btnActualizarWord").onclick = escribirTablaEnWord;
}

// ---------------------------------------------
// 2. LÓGICA DE REVISIONES (NUEVO)
// ---------------------------------------------

// Función para calcular la siguiente letra y descripción
function setNextLogic(type) {
    const lastRev = revisions.length > 0 ? revisions[revisions.length - 1].letra : null;
    let nextLetra = 'A';
    let nextDesc = '';
    const clientStd = document.getElementById('ddlEstandar').value;

    if (!lastRev) {
        // Primera revisión siempre A
        nextLetra = 'A';
        nextDesc = 'Revisión Interna Empresa de Ingeniería';
    } else {
        if (type === 'ITERATE' || type === 'UPDATE_TEXT') {
            // Si solo actualizamos texto, mantenemos la letra que el usuario haya puesto o calculamos la siguiente
            if (type === 'UPDATE_TEXT') {
                 // Si es update text, intentamos leer lo que ya hay, si está vacío calculamos
                 const currentInput = document.getElementById('txtRevLetra').value;
                 nextLetra = currentInput || String.fromCharCode(lastRev.charCodeAt(0) + 1);
            } else {
                // Cálculo: A->B, B->C...
                nextLetra = String.fromCharCode(lastRev.charCodeAt(0) + 1);
            }

            // Lógica de textos según cliente
            if (nextLetra === 'A') {
                nextDesc = 'Revisión Interna Empresa de Ingeniería';
            } else if (nextLetra === 'B') {
                nextDesc = (clientStd === 'CODELCO') ? 'Revisión de Codelco' : 'Revisión Cliente';
            } else {
                // C, D, E...
                nextDesc = (clientStd === 'CODELCO') ? 'Revisión de Codelco' : 'Revisión Cliente';
            }

        } else if (type === 'PHASE') {
            // Salto a Fase P
            if (lastRev < 'P') {
                nextLetra = 'P';
                nextDesc = 'Siguiente Fase'; // O "Apto para Construcción"
            } else {
                // Si ya estamos en P, sigue Q
                nextLetra = String.fromCharCode(lastRev.charCodeAt(0) + 1);
                nextDesc = 'Siguiente Fase';
            }
        }
    }

    // Rellenar inputs
    document.getElementById('txtRevLetra').value = nextLetra;
    document.getElementById('txtRevDesc').value = nextDesc;
    
    // Actualizar fecha a hoy siempre que se calcula nuevo
    establecerFechaHoyInput();
}

function addRevisionRow() {
    const letra = document.getElementById('txtRevLetra').value.toUpperCase();
    const fecha = document.getElementById('txtRevFecha').value;
    const desc = document.getElementById('txtRevDesc').value;

    if (!letra || !fecha) {
        mostrarMensaje("⚠️ Falta letra o fecha", "red");
        return;
    }

    // Agregar al array global
    revisions.push({ letra, fecha, desc });
    
    // Renderizar tabla HTML
    renderTable();
    
    // Calcular siguiente paso automáticamente para agilizar
    setNextLogic('ITERATE');
    mostrarMensaje("");
}

function renderTable() {
    const tbody = document.getElementById('tbodyRevisiones');
    tbody.innerHTML = '';

    // Invertimos para mostrar la más reciente arriba (stacking up)
    // OJO: Copiamos el array con [...revisions] para no invertir el original
    const displayRevisions = [...revisions].reverse(); 

    displayRevisions.forEach((rev, index) => {
        // Índice real en el array original (para poder borrar)
        const realIndex = revisions.length - 1 - index;
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><b>${rev.letra}</b></td>
            <td>${rev.fecha}</td>
            <td>${rev.desc}</td>
            <td style="text-align:right;">
                <span style="cursor:pointer; color:red; font-weight:bold;" onclick="deleteRev(${realIndex})">×</span>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

// Necesitamos exponer esta función al contexto global para que el onclick del HTML funcione
window.deleteRev = function(index) {
    revisions.splice(index, 1);
    renderTable();
};

// ---------------------------------------------
// 3. ESCRITURA EN WORD (TABLA)
// ---------------------------------------------

async function escribirTablaEnWord() {
    mostrarMensaje("⏳ Escribiendo en Word...", "blue");

    await Word.run(async (context) => {
        // Estrategia: Buscamos Content Controls dentro de la tabla de revisiones del Word.
        // Asumimos que la tabla tiene 3 columnas con tags: ccTablaRev, ccTablaFecha, ccTablaDesc
        
        // OJO: Si tu Word no tiene una tabla repetitiva automática, la forma más segura
        // es buscar los controles individuales (Rev A, Fecha A...) si son fijos.
        // Pero como describiste un sistema dinámico (filas que crecen hacia arriba),
        // Lo ideal es tener UNA tabla y agregar filas.
        
        // AQUÍ INTENTAREMOS LLENAR UNA TABLA EXISTENTE SI TIENE TAGS "ccRevRow"
        // Si no tienes tags de tabla, avísame y cambiamos a inserción simple.
        
        // MODO SIMPLE: Asumimos que tienes controles sueltos para la ÚLTIMA revisión
        // o MODO TABLA: Insertar filas. Usaremos MODO TABLA INVERSA (Insertar al inicio).

        // Vamos a buscar la tabla que contiene el control "ccRevision" (o uno nuevo "ccTablaAnchor")
        const controls = context.document.contentControls.getByTag("ccRevision"); // Usamos uno que ya tenías como ancla
        controls.load("parentTable");
        
        await context.sync();

        if (controls.items.length === 0) {
            mostrarMensaje("⚠️ No encontré el control 'ccRevision' en el documento para ubicar la tabla.", "red");
            return;
        }

        const table = controls.items[0].parentTable;
        table.load("rows");
        await context.sync();

        // Limpiar filas de datos antiguas (opcional, si quieres reescribir todo)
        // O simplemente agregar las nuevas. 
        // Para este ejemplo, vamos a ASUMIR que escribimos las revisiones que están en la lista
        // en las filas disponibles, o creamos nuevas.

        // Invertimos revisions para escribir de abajo hacia arriba (A abajo, P arriba)
        // O según tu formato visual. Tu imagen muestra:
        // P (Arriba)
        // B (Medio)
        // A (Abajo)
        // Encabezados
        
        // Entonces debemos escribir en ese orden visual en la tabla de Word.
        
        // Validamos si la tabla existe
        if (!table) {
            mostrarMensaje("❌ El control no está dentro de una tabla.", "red");
            return;
        }

        // --- LÓGICA DE ESCRITURA EN TABLA WORD ---
        // 1. Borramos contenido de datos (dejamos encabezado y footer si existen)
        // Esto es complejo si no sabemos índices exactos.
        // MEJOR ESTRATEGIA: Insertar datos en controles específicos si existen.
        
        // Plan B (Más seguro para tu caso):
        // Escribir SOLO la última revisión en los controles "ccRevision", "ccFecha", "ccEmision"
        // que ya tenías, para que al menos funcione lo básico.
        
        if (revisions.length > 0) {
            const ultimaRev = revisions[revisions.length - 1]; // La más nueva (ej: B)
            
            // Llenamos los controles individuales que ya tenías configurados en tu Word
            fillCc(context, "ccRevision", ultimaRev.letra);
            fillCc(context, "ccFecha", ultimaRev.fecha);
            fillCc(context, "ccEmision", ultimaRev.desc); // O ccEmitidoPara
            
            mostrarMensaje("✅ Última revisión (Rev " + ultimaRev.letra + ") actualizada.", "green");
        } else {
             mostrarMensaje("⚠️ Lista vacía. Agrega revisiones.", "orange");
        }

    }).catch(error => {
        console.error("Error:", error);
        mostrarMensaje("❌ Error escribiendo tabla: " + error.message, "red");
    });
}

// Helper para llenar controles simples
function fillCc(context, tagName, value) {
    const ccs = context.document.contentControls.getByTag(tagName);
    // No hacemos await sync aquí para velocidad, se hace en el batch
    context.sync().then(() => {
        if (ccs.items.length > 0) {
            ccs.items.forEach(cc => cc.insertText(value, "Replace"));
        }
    });
}


// ---------------------------------------------
// 4. LÓGICA DE AZURE Y DATOS PROYECTO (TU CÓDIGO VIEJO)
// ---------------------------------------------

async function cargarDocumentosDesdeAzure(idProyecto) {
    const ddl = document.getElementById("ddlDocumentos");
    if (!ddl) return;
    ddl.innerHTML = "<option>Cargando códigos...</option>";

    try {
        const response = await fetch(URL_POWER_AUTOMATE, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ "codigoProyecto": idProyecto }) 
        });

        if (!response.ok) throw new Error("Error Power Automate");
        const listaCruda = await response.json();

        // Filtro duplicados
        const documentosUnicos = [];
        const codigosVistos = new Set();

        listaCruda.forEach(doc => {
            const idUnico = doc.codFDA || doc.Nombre; 
            if (!codigosVistos.has(idUnico)) {
                codigosVistos.add(idUnico);
                documentosUnicos.push(doc);
            }
        });

        ddl.innerHTML = "";
        if (documentosUnicos.length === 0) {
            ddl.innerHTML = "<option>No se encontraron documentos</option>";
            return;
        }

        const optDef = document.createElement("option");
        optDef.text = "-- Seleccione un Código FDA --";
        optDef.value = "";
        ddl.appendChild(optDef);

        documentosUnicos.forEach(doc => {
            const opt = document.createElement("option");
            opt.text = doc.codFDA || doc.Nombre; 
            opt.value = doc.codFDA || ""; 
            opt.setAttribute("data-nombre", doc.Nombre || "");
            opt.setAttribute("data-cliente", doc.codCliente || "");
            ddl.appendChild(opt);
        });

    } catch (error) {
        console.error(error);
        ddl.innerHTML = "<option>Error al cargar lista</option>";
    }
}

async function insertarDocumentoSeleccionado() {
    const ddl = document.getElementById("ddlDocumentos");
    const opcionSeleccionada = ddl.options[ddl.selectedIndex];
    const codigoFDA = ddl.value;

    if (!codigoFDA) return;

    let nombreDoc = opcionSeleccionada.getAttribute("data-nombre");
    let codigoCliente = opcionSeleccionada.getAttribute("data-cliente");

    if (!codigoCliente || codigoCliente === "SIN-CODIGO") codigoCliente = "N/A";
    if (!nombreDoc) nombreDoc = "DOCUMENTO SIN NOMBRE";
    nombreDoc = nombreDoc.toUpperCase();

    await Word.run(async (context) => {
        const ccFDA = context.document.contentControls.getByTag("ccCodigoFDA");
        const ccCli = context.document.contentControls.getByTag("ccCodigoCliente");
        const ccNom = context.document.contentControls.getByTag("ccNombreDoc");

        ccFDA.load("items");
        ccCli.load("items");
        ccNom.load("items");
        await context.sync();

        if (ccFDA.items.length > 0) ccFDA.items.forEach(cc => cc.insertText(codigoFDA, "Replace"));
        if (ccCli.items.length > 0) ccCli.items.forEach(cc => cc.insertText(codigoCliente, "Replace"));
        if (ccNom.items.length > 0) ccNom.items.forEach(cc => cc.insertText(nombreDoc, "Replace"));

        await context.sync();
    }).catch(e => console.error(e));
}

// ---------------------------------------------
// 5. MEMORIA Y UTILES
// ---------------------------------------------

async function cargarDatosDeMemoria() {
    try {
        let esDocumentoValido = false;
        await Word.run(async (context) => {
            const ccCheck = context.document.contentControls.getByTag("ccCliente");
            ccCheck.load("items");
            await context.sync();
            if (ccCheck.items.length > 0) esDocumentoValido = true;
        });

        if (!esDocumentoValido) return;

        const jsonDatos = localStorage.getItem("FDA_ProyectoActual");
        if (jsonDatos) {
            const datos = JSON.parse(jsonDatos);
            setText("lblCliente",   datos.cliente);
            setText("lblDivision",  datos.division);
            setText("lblContrato",  datos.contrato);
            setText("lblProyecto",  datos.nombre);

            escribirDatosBaseEnWord(datos).catch(e => console.warn(e));

            if (datos.id) cargarDocumentosDesdeAzure(datos.id);
        }
    } catch (e) { console.error(e); }
}

async function escribirDatosBaseEnWord(datos) {
    await Word.run(async (context) => {
        const tagsMapa = [
            { t: "ccCliente",              v: datos.cliente },
            { t: "ccDivisión",             v: datos.division },
            { t: "ccContrato",             v: datos.contrato },
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

function establecerFechaHoyInput() {
    const txtFecha = document.getElementById("txtRevFecha");
    if (txtFecha) {
        const hoy = new Date();
        const dia = String(hoy.getDate()).padStart(2, '0');
        const mes = String(hoy.getMonth() + 1).padStart(2, '0');
        const anio = hoy.getFullYear();
        txtFecha.value = `${dia}/${mes}/${anio}`;
    }
}

function mostrarMensaje(texto, color = "black") {
    const el = document.getElementById("mensajeEstado");
    if (el) {
        el.textContent = texto;
        el.style.color = color;
    }
}