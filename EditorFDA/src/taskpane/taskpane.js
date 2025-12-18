/* global document, Office, Word, fetch, localStorage, window */

// 1. CONFIGURACIÓN (Global)
// -----------------------------------------------------------------------------
const URL_POWER_AUTOMATE = "https://defaultef8b3c00d87343e58b66d56c25f2bd.fe.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d88cc5b40d1b48bfa41f130960371fe1/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QAwT8H-2RLeYuIvy4ISgzt0sXfcBX0JGvjjR_3l1V_Y"; 

// Variable global para almacenar el historial de revisiones (A, B, C...)
let revisions = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Office.js listo. Iniciando Taskpane...");
        asignarEventos();
        cargarDatosDeMemoria();
        establecerFechaHoyInput();
        setNextLogic('INIT');
    }
});

function asignarEventos() {
    const ddlDocs = document.getElementById("ddlDocumentos");
    if (ddlDocs) ddlDocs.onchange = insertarDocumentoSeleccionado;

    // --- EVENTOS DE REVISIONES ---
    document.getElementById("btnIterar").onclick = () => setNextLogic('ITERATE');
    document.getElementById("btnFase").onclick = () => setNextLogic('PHASE');
    document.getElementById("ddlEstandar").onchange = () => setNextLogic('UPDATE_TEXT');
    document.getElementById("btnAgregarAlista").onclick = addRevisionRow;
    
    // EL BOTÓN IMPORTANTE
    document.getElementById("btnActualizarWord").onclick = escribirTablaEnWord;
}

// ---------------------------------------------
// 2. LÓGICA DE REVISIONES (UI & VALIDACIÓN)
// ---------------------------------------------

function setNextLogic(type) {
    const lastRev = revisions.length > 0 ? revisions[revisions.length - 1].letra : null;
    let nextLetra = 'A';
    let nextDesc = '';
    const clientStd = document.getElementById('ddlEstandar').value;

    if (!lastRev) {
        nextLetra = 'A';
        nextDesc = 'Revisión Interna Empresa de Ingeniería';
    } else {
        if (type === 'ITERATE' || type === 'UPDATE_TEXT') {
            if (type === 'UPDATE_TEXT') {
                 const currentInput = document.getElementById('txtRevLetra').value;
                 nextLetra = currentInput || String.fromCharCode(lastRev.charCodeAt(0) + 1);
            } else {
                nextLetra = String.fromCharCode(lastRev.charCodeAt(0) + 1);
            }

            if (nextLetra === 'A') {
                nextDesc = 'Revisión Interna Empresa de Ingeniería';
            } else if (nextLetra === 'B') {
                nextDesc = (clientStd === 'CODELCO') ? 'Revisión de Codelco' : 'Revisión Cliente';
            } else {
                nextDesc = (clientStd === 'CODELCO') ? 'Revisión de Codelco' : 'Revisión Cliente';
            }

        } else if (type === 'PHASE') {
            if (lastRev < 'P') {
                nextLetra = 'P';
                nextDesc = 'Siguiente Fase'; 
            } else {
                nextLetra = String.fromCharCode(lastRev.charCodeAt(0) + 1);
                nextDesc = 'Siguiente Fase';
            }
        }
    }
    document.getElementById('txtRevLetra').value = nextLetra;
    document.getElementById('txtRevDesc').value = nextDesc;
    establecerFechaHoyInput();
}

function addRevisionRow() {
    const letra = document.getElementById('txtRevLetra').value.toUpperCase().trim();
    const fecha = document.getElementById('txtRevFecha').value;
    const desc = document.getElementById('txtRevDesc').value.trim();

    if (!letra || !fecha) {
        mostrarMensaje("⚠️ Falta letra o fecha.", "orange");
        return;
    }

    const existe = revisions.some(r => r.letra === letra);
    if (existe) {
        mostrarMensaje(`⛔ Error: La revisión "${letra}" ya existe en la lista.`, "red");
        return;
    }

    revisions.push({ letra, fecha, desc });
    
    // Ordenamos siempre alfabéticamente (A, B, C... P)
    revisions.sort((a, b) => {
        if (a.letra === b.letra) return 0;
        return a.letra > b.letra ? 1 : -1;
    });

    renderTable();
    setNextLogic('ITERATE');
    mostrarMensaje("");
}

function renderTable() {
    const tbody = document.getElementById('tbodyRevisiones');
    tbody.innerHTML = '';
    // Invertimos para visualización (stack up)
    const displayRevisions = [...revisions].reverse(); 

    displayRevisions.forEach((rev, index) => {
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

window.deleteRev = function(index) {
    revisions.splice(index, 1);
    renderTable();
};

// ---------------------------------------------
// 3. ESCRITURA EN WORD (INTELIGENTE: ACTUALIZA O INSERTA)
// ---------------------------------------------

async function escribirTablaEnWord() {
    mostrarMensaje("⏳ Analizando, reciclando y escribiendo...", "blue");

    if (revisions.length === 0) {
        mostrarMensaje("⚠️ Lista vacía. Agrega revisiones primero.", "orange");
        return;
    }

    await Word.run(async (context) => {
        // 1. OBTENER TABLA
        const contentControls = context.document.contentControls.getByTag("ccTablaRevisiones");
        contentControls.load("items/tables/items");
        await context.sync();

        if (contentControls.items.length === 0 || contentControls.items[0].tables.items.length === 0) {
            mostrarMensaje("❌ No encontré la tabla 'ccTablaRevisiones'.", "red");
            return;
        }

        const tablaWord = contentControls.items[0].tables.items[0];
        const filasWord = tablaWord.rows;
        
        // Cargamos valores para leer y body para escribir
        filasWord.load("items/cells/items/value, items/cells/items/body");
        await context.sync();

        // 2. PREPARACIÓN DE MAPAS
        // Mapa de lo que QUEREMOS tener (Tu lista del panel)
        let mapaDeseado = new Map();
        revisions.forEach(r => mapaDeseado.set(r.letra, r));

        // Array para guardar los índices de filas que podemos REUTILIZAR (Vacías u Obsoletas)
        let slotsDisponibles = [];

        // 3. PRIMERA PASADA: ACTUALIZAR COINCIDENCIAS Y ENCONTRAR HUECOS
        // Iteramos filas (menos el encabezado final)
        for (let i = 0; i < filasWord.items.length - 1; i++) {
            let fila = filasWord.items[i];
            
            // Protección contra filas rotas
            if (fila.cells.items.length < 3) continue;

            let valorCelda = fila.cells.items[0].value;
            let letraEnWord = valorCelda ? valorCelda.trim().toUpperCase() : "";

            if (letraEnWord === "") {
                // CASO 1: LA FILA ESTÁ VACÍA -> Es un slot disponible
                slotsDisponibles.push(fila);
            } 
            else if (mapaDeseado.has(letraEnWord)) {
                // CASO 2: COINCIDENCIA (Ej: A=A) -> Actualizamos datos ahí mismo
                let datos = mapaDeseado.get(letraEnWord);
                console.log(`Actualizando fila existente: ${letraEnWord}`);
                
                fila.cells.items[1].body.insertText(datos.fecha, "Replace");
                fila.cells.items[2].body.insertText(datos.desc, "Replace");
                
                // Quitamos del mapa porque ya está lista
                mapaDeseado.delete(letraEnWord);
            } 
            else {
                // CASO 3: OBSOLETA (Ej: P, o una C vieja que ya no queremos) 
                // -> La limpiamos y la marcamos como disponible para usarla después
                console.log(`Reciclando fila obsoleta: ${letraEnWord}`);
                
                // Borramos contenido visualmente
                fila.cells.items[0].body.insertText("", "Replace");
                fila.cells.items[1].body.insertText("", "Replace");
                fila.cells.items[2].body.insertText("", "Replace");
                
                slotsDisponibles.push(fila);
            }
        }

        // 4. SEGUNDA PASADA: LLENAR SLOTS VACÍOS CON LAS NUEVAS (Ej: C, D)
        // Lo que quedó en 'mapaDeseado' son las revisiones nuevas que no estaban en Word.
        // Debemos insertarlas. Primero intentamos usar los slots vacíos.

        // Convertimos el mapa a array y lo INVERTIMOS para priorizar las más nuevas arriba (Stack)
        // Revisions original: [A, B, C, D]
        // Si A y B ya se actualizaron, en el mapa quedan C y D.
        // Queremos escribir D arriba y C abajo (si hubiera espacio).
        
        // Para respetar el orden de pila (Start), las procesamos desde la más nueva (D) hacia atrás.
        let pendientes = [...revisions].reverse().filter(r => mapaDeseado.has(r.letra));

        // Slots disponibles suelen estar de arriba a abajo (índice 0, 1...).
        // Si tenemos slots en fila 0 y fila 1.
        // Y tenemos pendientes D y C.
        // D debería ir en fila 0. C en fila 1.
        // Así que simplemente iteramos los slots en orden.

        let filasNuevasParaCrear = [];

        for (let rev of pendientes) {
            if (slotsDisponibles.length > 0) {
                // --- USAR SLOT RECICLADO ---
                // Tomamos el primer slot disponible (el de más arriba visualmente)
                let slot = slotsDisponibles.shift(); // Saca el primero del array
                
                console.log(`Escribiendo ${rev.letra} en slot reciclado.`);
                slot.cells.items[0].body.insertText(rev.letra, "Replace");
                slot.cells.items[1].body.insertText(rev.fecha, "Replace");
                slot.cells.items[2].body.insertText(rev.desc, "Replace");
            } else {
                // --- NO HAY SLOTS -> AGREGAR A LA COLA DE CREACIÓN ---
                // Si se acabaron los huecos, guardamos para crear fila nueva al final
                filasNuevasParaCrear.push([rev.letra, rev.fecha, rev.desc]);
            }
        }

        // 5. TERCERA PASADA: CREAR FILAS SI FALTARON SLOTS
        if (filasNuevasParaCrear.length > 0) {
            console.log(`Creando ${filasNuevasParaCrear.length} filas nuevas reales.`);
            // 'Start' las pone arriba del todo
            tablaWord.addRows("Start", filasNuevasParaCrear.length, filasNuevasParaCrear);
        }

        await context.sync();
        mostrarMensaje("✅ Tabla actualizada correctamente.", "green");

    }).catch(error => {
        console.error("Error Word:", error);
        mostrarMensaje("❌ Error: " + error.message, "red");
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
        ccFDA.load("items"); ccCli.load("items"); ccNom.load("items");
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