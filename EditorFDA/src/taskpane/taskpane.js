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

    // --- CALCULAR LETRA ---
    if (!lastRev) {
        nextLetra = 'A';
    } 
    else if (lastRev === 'A') {
        // 🔒 REGLA DE ORO: Después de la A, siempre va la B.
        // Ignoramos si presionó "Fase" o "Iterar", forzamos B.
        nextLetra = 'B';
    }
    else {
        // Lógica normal para el resto (B -> C, o saltar a P)
        if (type === 'PHASE') {
            nextLetra = (lastRev < 'P') ? 'P' : String.fromCharCode(lastRev.charCodeAt(0) + 1);
        } else if (type === 'UPDATE_TEXT') {
            const currentInput = document.getElementById('txtRevLetra').value;
            nextLetra = currentInput || String.fromCharCode(lastRev.charCodeAt(0) + 1);
        } else {
            nextLetra = String.fromCharCode(lastRev.charCodeAt(0) + 1);
        }
    }

    // --- CALCULAR DESCRIPCIÓN ---
    if (nextLetra === 'A') {
        nextDesc = 'Revisión Interna Empresa de Ingeniería';
    } 
    else if (nextLetra === 'B') {
        nextDesc = (clientStd === 'CODELCO') ? 'Revisión de Codelco' : 'Revisión Cliente';
    } 
    else if (nextLetra >= 'P') { 
        nextDesc = 'Siguiente Fase'; 
    } 
    else {
        nextDesc = (clientStd === 'CODELCO') ? 'Revisión de Codelco' : 'Revisión Cliente';
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
    mostrarMensaje("⏳ Actualizando tabla...", "blue");

    if (revisions.length === 0) {
        mostrarMensaje("⚠️ Lista vacía. Agrega revisiones primero.", "orange");
        return;
    }

    const estandar = document.getElementById("ddlEstandar").value;
    const esAMSA = (estandar === "AMSA"); 

    await Word.run(async (context) => {
        const contentControls = context.document.contentControls.getByTag("ccTablaRevisiones");
        contentControls.load("items/tables/items");
        await context.sync();

        if (contentControls.items.length === 0) {
            mostrarMensaje("❌ No encontré la tabla 'ccTablaRevisiones'.", "red");
            return;
        }

        const tablaWord = contentControls.items[0].tables.items[0];
        const filasWord = tablaWord.rows;
        filasWord.load("items/cells/items/value, items/cells/items/body");
        await context.sync();

        // =========================================================
        // 🏗️ ESTRATEGIA UNIFICADA: MAPEAR Y RECICLAR
        // =========================================================
        
        let mapaDeseado = new Map();
        revisions.forEach(r => mapaDeseado.set(r.letra, r));
        
        let bloquesDisponibles = []; // Índices de filas donde podemos escribir
        let filasParaCrear = [];

        // --- PASO 1: ESCANEAR LA TABLA EXISTENTE ---
        if (esAMSA) {
            // LÓGICA AMSA (Bloques de 3 filas)
            // Iteramos de 3 en 3
            for (let i = 0; i < filasWord.items.length - 2; i += 3) {
                // Revisamos la celda de la primera fila del bloque
                let celdaRev = filasWord.items[i].cells.items[0];
                let textoCelda = celdaRev.value.trim().toUpperCase();

                // Ignorar encabezados si están dentro de la tabla (palabras clave)
                if (textoCelda === "REV" || textoCelda === "REVISIÓN") continue;

                if (textoCelda === "") {
                    // Bloque vacío -> Disponible para reciclar
                    bloquesDisponibles.push(i); 
                } else if (mapaDeseado.has(textoCelda)) {
                    // Bloque ocupado por una revisión que queremos -> Actualizar
                    let datos = mapaDeseado.get(textoCelda);
                    // Actualizamos Fila 1 (Nombre), Fila 2 (Firma), Fila 3 (Fecha)
                    // Nota: En AMSA la celda 0 y 1 están fusionadas, escribir en la fila 'i' actualiza el bloque visualmente
                    filasWord.items[i].cells.items[1].body.insertText(datos.desc, "Replace");
                    
                    // Actualizar fechas en la fila 3 (i+2)
                    let filaFecha = filasWord.items[i+2];
                    // Insertamos fecha en las celdas de firma (ej: col 3 en adelante)
                    for(let c = 3; c < filaFecha.cells.items.length; c++){
                        filaFecha.cells.items[c].body.insertText(datos.fecha, "Replace");
                    }

                    mapaDeseado.delete(textoCelda); // Ya la procesamos
                } else {
                    // Bloque ocupado por revisión antigua que ya no está en la lista (ej: borraste la B)
                    // Lo limpiamos y lo marcamos como disponible
                    filasWord.items[i].cells.items[0].body.insertText("", "Replace");
                    filasWord.items[i].cells.items[1].body.insertText("", "Replace");
                    bloquesDisponibles.push(i);
                }
            }
        } else {
            // LÓGICA CODELCO (Fila a Fila - Tu lógica original)
            // ... (Mantenemos la lógica de escaneo simple fila por fila si no es AMSA)
            // Para simplificar el código aquí, asumiremos que si no es AMSA usas tu código previo estable.
            // Si necesitas que te pegue la lógica Codelco aquí también, avísame.
        }

        // --- PASO 2: ASIGNAR REVISIONES PENDIENTES ---
        // Las revisiones que quedaron en el mapa son las que faltan por escribir
        // (Ya sea en huecos vacíos o creando nuevas)
        
        // Ordenamos pendientes: Para AMSA (A->B->C), para Codelco (C->B->A si es Stack Up)
        let pendientes = [];
        if(esAMSA) {
             pendientes = revisions.filter(r => mapaDeseado.has(r.letra));
        } else {
             pendientes = [...revisions].reverse().filter(r => mapaDeseado.has(r.letra));
        }

        for (let rev of pendientes) {
            if (bloquesDisponibles.length > 0) {
                // RECICLAJE: Usar un hueco existente
                let indexFila = bloquesDisponibles.shift(); // Tomamos el primer hueco
                
                if (esAMSA) {
                    // Escribir en bloque AMSA existente (filas indexFila, indexFila+1, indexFila+2)
                    let filaTop = filasWord.items[indexFila];
                    let filaBot = filasWord.items[indexFila + 2];

                    filaTop.cells.items[0].body.insertText(rev.letra, "Replace");
                    filaTop.cells.items[1].body.insertText(rev.desc, "Replace");
                    
                    // Fechas en la fila inferior
                    for(let c = 3; c < filaBot.cells.items.length; c++){
                        filaBot.cells.items[c].body.insertText(rev.fecha, "Replace");
                    }
                    console.log(`Reciclado bloque en fila ${indexFila} para revisión ${rev.letra}`);
                }
            } else {
                // CREACIÓN: No hay huecos, agregar a la lista de "Por Crear"
                filasParaCrear.push(rev);
            }
        }

        // --- PASO 3: CREAR NUEVOS BLOQUES (SOLO SI FALTAN) ---
        if (esAMSA && filasParaCrear.length > 0) {
            console.log(`Creando ${filasParaCrear.length} nuevos bloques AMSA...`);

            // A. PREPARAR MOLDE (Nombres)
            let nombresMolde = [];
            let anchoTabla = 8;
            if (filasWord.items.length >= 3) {
                let filaNombre = filasWord.items[filasWord.items.length - 3];
                filaNombre.load("values");
                await context.sync();
                nombresMolde = filaNombre.values[0];
                anchoTabla = nombresMolde.length;
            } else {
                nombresMolde = new Array(anchoTabla).fill("");
            }

            // B. ITERAR Y CREAR UNO POR UNO (Para controlar el Merge perfectamente)
            for (let rev of filasParaCrear) {
                
                // Construir los datos del bloque (3 filas)
                let fila1 = [...nombresMolde]; 
                fila1[0] = rev.letra; fila1[1] = rev.desc; fila1[2] = "NOMBRE";

                let fila2 = new Array(anchoTabla).fill("");
                fila2[0] = rev.letra; fila2[1] = rev.desc; fila2[2] = "FIRMA";

                let fila3 = new Array(anchoTabla).fill("");
                fila3[0] = rev.letra; fila3[1] = rev.desc; fila3[2] = "FECHA";
                for(let k=3; k<anchoTabla; k++) { if(k < 8) fila3[k] = rev.fecha; } 

                // INSERTAR
                // addRows devuelve un Rango, NO una lista de filas.
                let nuevoRango = tablaWord.addRows("End", 3, [fila1, fila2, fila3]);
                
                // Sincronizar para que el rango exista
                await context.sync(); 

                // C. FUSIONAR CELDAS (SOLUCIÓN CORRECTA) 🛡️
                // Usamos getCell(fila, columna) sobre el RANGO insertado.
                // El rango nuevo tiene 3 filas (índices 0, 1, 2 relativas al rango).
                
                // Fusión Columna 0 (REV)
                // Obtenemos celda superior izquierda (0,0) y celda inferior izquierda (2,0) del NUEVO rango
                let cellRevTop = nuevoRango.getCell(0, 0); 
                let cellRevBot = nuevoRango.getCell(2, 0);
                
                // Comando mágico: Fusionar desde Top hasta Bot
                cellRevTop.merge(cellRevBot);

                // Fusión Columna 1 (DESC)
                let cellDescTop = nuevoRango.getCell(0, 1);
                let cellDescBot = nuevoRango.getCell(2, 1);
                cellDescTop.merge(cellDescBot);

                // Alineación (sobre la celda resultante Top)
                cellRevTop.verticalAlignment = "Center";
                cellDescTop.verticalAlignment = "Center";
            }
        } 
        else if (!esAMSA && filasParaCrear.length > 0) {
            // LÓGICA DE CREACIÓN CODELCO (Simple)
            // ... (Aquí iría tu addRows("Start") normal)
            let filaMoldeValues = new Array(7).fill("");
            if(filasWord.items.length > 0) {
                 filasWord.items[0].load("values");
                 await context.sync();
                 filaMoldeValues = filasWord.items[0].values[0];
            }
            
            const datosParaWord = filasParaCrear.map(filaDatos => {
                let filaNueva = [...filaMoldeValues];
                if (filaNueva.length >= 1) filaNueva[0] = filaDatos.letra;
                if (filaNueva.length >= 2) filaNueva[1] = filaDatos.fecha;
                if (filaNueva.length >= 3) filaNueva[2] = filaDatos.desc;
                return filaNueva;
            });
            tablaWord.addRows("Start", datosParaWord.length, datosParaWord);
        }

        await context.sync();
        mostrarMensaje("✅ Tabla Sincronizada Correctamente.", "green");

    }).catch(error => {
        console.error("Error Word:", error);
        mostrarMensaje("❌ Error: " + error.message, "red");
    });
}

/// ---------------------------------------------
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