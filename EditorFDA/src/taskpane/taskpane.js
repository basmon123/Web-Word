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

// ---------------------------------------------
// 3. ESCRITURA EN WORD (TABLA) - CORREGIDO
// ---------------------------------------------

async function escribirTablaEnWord() {
    mostrarMensaje("⏳ Escribiendo en Word...", "blue");

    // Verificar si hay datos que escribir
    if (revisions.length === 0) {
        mostrarMensaje("⚠️ La lista de revisiones está vacía. Agrega una primero.", "orange");
        return;
    }

    // Tomamos la ÚLTIMA revisión de la lista (la más reciente)
    const ultimaRev = revisions[revisions.length - 1];

    await Word.run(async (context) => {
        // 1. OBTENER REFERENCIAS A LOS CONTROLES POR SU TAG
        // Asegúrate que en tu Word los controles tengan estos Tags exactos:
        const ccRevs = context.document.contentControls.getByTag("ccRevision");
        const ccFechas = context.document.contentControls.getByTag("ccFecha");
        const ccEmisiones = context.document.contentControls.getByTag("ccEmision"); // O ccEmitidoPara

        // 2. CARGAR LA PROPIEDAD 'items' (¡Esto es lo que faltaba!)
        // Sin esto, Word no sabe cuántos controles hay ni cuáles son.
        ccRevs.load("items");
        ccFechas.load("items");
        ccEmisiones.load("items");

        // 3. SINCRONIZAR (Traer los objetos de Word a la memoria de JS)
        await context.sync();

        // 4. ESCRIBIR DATOS (Ahora sí es seguro leer .items)
        let controlesEncontrados = false;

        // Escribir Letra Revisión
        if (ccRevs.items.length > 0) {
            ccRevs.items.forEach(cc => cc.insertText(ultimaRev.letra, "Replace"));
            controlesEncontrados = true;
        }

        // Escribir Fecha
        if (ccFechas.items.length > 0) {
            ccFechas.items.forEach(cc => cc.insertText(ultimaRev.fecha, "Replace"));
            controlesEncontrados = true;
        }

        // Escribir Descripción
        if (ccEmisiones.items.length > 0) {
            ccEmisiones.items.forEach(cc => cc.insertText(ultimaRev.desc, "Replace"));
            controlesEncontrados = true;
        }

        if (!controlesEncontrados) {
            mostrarMensaje("⚠️ No encontré los controles (tags: ccRevision, ccFecha, ccEmision) en el Word.", "orange");
            return;
        }

        // 5. SINCRONIZAR FINAL (Enviar los textos nuevos a Word)
        await context.sync();
        
        mostrarMensaje(`✅ Revisión ${ultimaRev.letra} actualizada en el documento.`, "green");

    }).catch(error => {
        console.error("Error:", error);
        mostrarMensaje("❌ Error: " + error.message, "red");
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