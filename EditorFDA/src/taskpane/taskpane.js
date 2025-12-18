/* global document, Office, Word, fetch, localStorage, window */

// 1. CONFIGURACIÓN (Global)
// -----------------------------------------------------------------------------
const URL_POWER_AUTOMATE = "https://defaultef8b3c00d87343e58b66d56c25f2bd.fe.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d88cc5b40d1b48bfa41f130960371fe1/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QAwT8H-2RLeYuIvy4ISgzt0sXfcBX0JGvjjR_3l1V_Y"; 

// Variable global para almacenar el historial de revisiones
let revisions = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Office.js listo. Iniciando Taskpane...");

        asignarEventos();
        cargarDatosDeMemoria();
        
        // Inicializar fecha actual en el input
        establecerFechaHoyInput();
        
        // Inicializar lógica (sugerir Rev A)
        setNextLogic('INIT');
    }
});

function asignarEventos() {
    // Evento Dropdown Documentos (Azure)
    const ddlDocs = document.getElementById("ddlDocumentos");
    if (ddlDocs) ddlDocs.onchange = insertarDocumentoSeleccionado;

    // --- EVENTOS DE REVISIONES ---
    
    // 1. Botones de Lógica (Iterar / Fase)
    document.getElementById("btnIterar").onclick = () => setNextLogic('ITERATE');
    document.getElementById("btnFase").onclick = () => setNextLogic('PHASE');
    
    // 2. Cambio de Estándar (Codelco vs AMSA) actualiza textos sugeridos
    document.getElementById("ddlEstandar").onchange = () => setNextLogic('UPDATE_TEXT');

    // 3. Botón "Insertar en Lista Arriba" (Panel visual)
    document.getElementById("btnAgregarAlista").onclick = addRevisionRow;

    // 4. Botón "Actualizar Tabla en Documento" (Escribir en Word)
    document.getElementById("btnActualizarWord").onclick = escribirTablaEnWord;
}

// ---------------------------------------------
// 2. LÓGICA DE REVISIONES (UI & VALIDACIÓN)
// ---------------------------------------------

// Función para calcular automáticamente la siguiente letra y descripción
function setNextLogic(type) {
    // Obtenemos la última letra agregada para saber qué sigue
    const lastRev = revisions.length > 0 ? revisions[revisions.length - 1].letra : null;
    let nextLetra = 'A';
    let nextDesc = '';
    const clientStd = document.getElementById('ddlEstandar').value;

    if (!lastRev) {
        // Si la lista está vacía, partimos con A
        nextLetra = 'A';
        nextDesc = 'Revisión Interna Empresa de Ingeniería';
    } else {
        if (type === 'ITERATE' || type === 'UPDATE_TEXT') {
            // Si es actualizar texto, mantenemos la letra actual del input si existe
            if (type === 'UPDATE_TEXT') {
                 const currentInput = document.getElementById('txtRevLetra').value;
                 nextLetra = currentInput || String.fromCharCode(lastRev.charCodeAt(0) + 1);
            } else {
                // Cálculo matemático: A->B, B->C...
                nextLetra = String.fromCharCode(lastRev.charCodeAt(0) + 1);
            }

            // Lógica de Textos según cliente
            if (nextLetra === 'A') {
                nextDesc = 'Revisión Interna Empresa de Ingeniería';
            } else if (nextLetra === 'B') {
                nextDesc = (clientStd === 'CODELCO') ? 'Revisión de Codelco' : 'Revisión Cliente';
            } else {
                // Para C, D, E...
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

    // Rellenar inputs visuales
    document.getElementById('txtRevLetra').value = nextLetra;
    document.getElementById('txtRevDesc').value = nextDesc;
    
    // Actualizar fecha a hoy
    establecerFechaHoyInput();
}

// Agregar fila a la lista visual del Taskpane
function addRevisionRow() {
    const letra = document.getElementById('txtRevLetra').value.toUpperCase().trim();
    const fecha = document.getElementById('txtRevFecha').value;
    const desc = document.getElementById('txtRevDesc').value.trim();

    // 1. VALIDACIÓN BÁSICA
    if (!letra || !fecha) {
        mostrarMensaje("⚠️ Falta letra o fecha.", "orange");
        return;
    }

    // 2. VALIDACIÓN DE DUPLICADOS (Evita dos 'B' o dos 'A')
    const existe = revisions.some(r => r.letra === letra);
    if (existe) {
        mostrarMensaje(`⛔ Error: La revisión "${letra}" ya existe en la lista.`, "red");
        return;
    }

    // 3. AGREGAR Y ORDENAR
    revisions.push({ letra, fecha, desc });
    
    // Ordenamos siempre alfabéticamente (A, B, C... P) para mantener consistencia
    revisions.sort((a, b) => {
        if (a.letra === b.letra) return 0;
        return a.letra > b.letra ? 1 : -1;
    });

    renderTable(); // Dibujar tabla HTML
    
    // Calcular siguiente paso automáticamente
    setNextLogic('ITERATE');
    mostrarMensaje("");
}

// Dibujar la tabla HTML en el panel lateral
function renderTable() {
    const tbody = document.getElementById('tbodyRevisiones');
    tbody.innerHTML = '';

    // Invertimos para mostrar la más reciente ARRIBA (Pila visual: C sobre B sobre A)
    const displayRevisions = [...revisions].reverse(); 

    displayRevisions.forEach((rev, index) => {
        // Calculamos el índice real para poder borrar correctamente
        const realIndex = revisions.length - 1 - index;
        
        const tr = document.createElement('tr');
        // Usamos estilos inline básicos compatibles o clases si prefieres
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

// Función global para borrar filas del panel
window.deleteRev = function(index) {
    revisions.splice(index, 1);
    renderTable();
};

// ---------------------------------------------
// 3. ESCRITURA EN WORD (MODO PILA COMPLETA)
// ---------------------------------------------

async function escribirTablaEnWord() {
    mostrarMensaje("⏳ Insertando historial en Word...", "blue");

    if (revisions.length === 0) {
        mostrarMensaje("⚠️ Lista vacía. Agrega revisiones en el panel primero.", "orange");
        return;
    }

    await Word.run(async (context) => {
        // 1. Buscamos el control contenedor de la tabla
        // IMPORTANTE: El tag "ccTablaRevisiones" debe estar en la FILA DE DATOS (la vacía sobre el encabezado)
        const contentControls = context.document.contentControls.getByTag("ccTablaRevisiones");
        contentControls.load("items");
        
        await context.sync();

        if (contentControls.items.length === 0) {
            mostrarMensaje("❌ No encontré el tag 'ccTablaRevisiones'. Verifica tu Word.", "red");
            return;
        }

        // 2. Obtenemos la tabla dentro del control
        const control = contentControls.items[0];
        const tablas = control.tables;
        tablas.load("items");
        
        await context.sync();

        if (tablas.items.length === 0) {
            mostrarMensaje("❌ El control no contiene una tabla válida.", "red");
            return;
        }

        const tablaWord = tablas.items[0];

        // 3. PREPARAR DATOS (INVERTIDOS)
        // Queremos que en el Word quede:
        // C (Arriba)
        // B
        // A (Abajo)
        // Encabezado (Fijo)
        // Por lo tanto, invertimos el array y mandamos el bloque completo.
        
        const datosParaWord = [...revisions].reverse().map(rev => {
            return [rev.letra, rev.fecha, rev.desc];
        });

        // 4. INSERTAR BLOQUE COMPLETO AL INICIO ('Start')
        // Esto empujará lo que ya exista hacia abajo.
        // Si tenías filas viejas o vacías, quedarán debajo de las nuevas.
        // El usuario solo debe borrar las sobrantes manualmente si es necesario.
        
        tablaWord.addRows("Start", datosParaWord.length, datosParaWord);
        
        await context.sync();
        
        mostrarMensaje(`✅ Se insertaron ${datosParaWord.length} revisiones correctamente.`, "green");

    }).catch(error => {
        console.error("Error Word:", error);
        mostrarMensaje("❌ Error al escribir: " + error.message, "red");
    });
}


// ---------------------------------------------
// 4. LÓGICA DE AZURE Y DATOS PROYECTO (ORIGINAL)
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
        //// Usamos formato texto dd/mm/aaaa
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