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
    mostrarMensaje("⏳ Procesando tabla...", "blue");

    if (revisions.length === 0) {
        mostrarMensaje("⚠️ Lista vacía. Agrega revisiones primero.", "orange");
        return;
    }

    const estandar = document.getElementById("ddlEstandar").value;
    const esAMSA = (estandar === "AMSA"); 

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
        filasWord.load("items/cells/items/value, items/cells/items/body");
        await context.sync();

        // 2. PREPARACIÓN DE DATOS
        let mapaDeseado = new Map();
        revisions.forEach(r => mapaDeseado.set(r.letra, r));
        
        // --- LISTA BLINDADA PARA AMSA Y CODELCO ---
        // Basada en tus imágenes: Protege "N° INTERNO", "POR", "REVISIONES", "REV", etc.
        const palabrasProtegidas = [
            "REVISIÓN", "REVISION", "REV", "REV.", // "REV" sin punto es vital para AMSA
            "REVISIONES", // Vital para encabezado AMSA
            "FECHA", 
            "EMITIDO", 
            "DESCRIPCIÓN", "DESCRIPCION", 
            "PROYECTO", 
            "TÍTULO", "TITULO", // Vital para encabezado AMSA
            "FDA", "CENTINELA", // Nombres de empresas en encabezados
            "APROBÓ", "APROBO", 
            "PREPARÓ", "PREPARO", "POR", // "POR" vital para AMSA
            "REVISÓ", "REVISO", 
            "CLIENTE", 
            "N°", "N.", "NO.", "INTERNO", // "INTERNO" protege "N° INTERNO"
            "NOMBRE", "FIRMA" // Etiquetas laterales de AMSA
        ];

        // =========================================================
        // 🏗️ MODO AMSA (STACK DOWN)
        // =========================================================
        if (esAMSA) {
            console.log("🔵 MODO AMSA ACTIVADO");

            let bloquesDisponibles = [];
            let letrasYaEnTabla = new Set();

            // Iteramos de 3 en 3 (saltando encabezados protegidos)
            // IMPORTANTE: Ajustamos el loop para no romper si la tabla no es múltiplo de 3 exacto
            // debido a los encabezados superiores.
            
            // Paso A: Escanear toda la tabla fila por fila para identificar qué es qué
            for (let i = 0; i < filasWord.items.length; i++) {
                let celdaRev = filasWord.items[i].cells.items[0];
                let texto = celdaRev.value.trim().toUpperCase();

                // 1. CHEQUEO DE PROTECCIÓN (ENCABEZADOS)
                // Si la celda contiene alguna palabra protegida, SALTAMOS esa fila.
                // Usamos 'some' para ver si incluye alguna palabra clave.
                if (palabrasProtegidas.some(p => texto.includes(p))) {
                    continue; 
                }

                // 2. DETECCIÓN DE BLOQUES DE DATOS
                // Si llegamos aquí, NO es un encabezado. Es una fila de datos o vacía.
                // En AMSA, los datos vienen en bloques de 3. Verificamos si es el inicio de un bloque.
                // Un bloque válido tiene que tener espacio para 3 filas (i, i+1, i+2)
                if (i + 2 < filasWord.items.length) {
                    // Verificamos si esta fila 'i' es el inicio (letras A, B, C o vacío)
                    // y que la fila 'i+1' tenga "FIRMA" (o vacío) en la columna lateral (índice 2)
                    // Esto ayuda a confirmar que estamos en el "grid" correcto.
                    
                    // Cargamos valor de la etiqueta lateral (columna 2) para confirmar estructura AMSA
                    let etiquetaLateral = filasWord.items[i].cells.items[2].value.trim().toUpperCase(); // Debería ser "NOMBRE" o vacío
                    let etiquetaSiguiente = filasWord.items[i+1].cells.items[2].value.trim().toUpperCase(); // Debería ser "FIRMA"

                    // Si parece un bloque AMSA (o está vacío y disponible)
                    if (texto === "" && (etiquetaLateral === "" || etiquetaLateral === "NOMBRE")) {
                        bloquesDisponibles.push(i);
                        i += 2; // Saltamos las otras 2 filas del bloque
                    } 
                    else if (mapaDeseado.has(texto)) {
                        // ES UNA REVISIÓN EXISTENTE (A, B...)
                        let datos = mapaDeseado.get(texto);
                        
                        // Fila 1 (Nombre)
                        filasWord.items[i].cells.items[1].body.insertText(datos.desc, "Replace");
                        // Fila 3 (Fechas)
                        let filaFechas = filasWord.items[i + 2];
                        for(let c=3; c < filaFechas.cells.items.length; c++) {
                             // Solo escribimos si la celda no está vacía (asumiendo que tiene borde/formato)
                             // O simplemente escribimos sin miedo:
                             filaFechas.cells.items[c].body.insertText(datos.fecha, "Replace");
                        }
                        letrasYaEnTabla.add(texto);
                        i += 2; // Saltamos el bloque ya procesado
                    }
                    // Ojo: Si hay una fila suelta que no es bloque ni encabezado, el loop sigue normal.
                }
            }

            // B. DETERMINAR PENDIENTES
            let pendientes = revisions.filter(r => !letrasYaEnTabla.has(r.letra));

            // C. RECICLAR HUECOS
            while (bloquesDisponibles.length > 0 && pendientes.length > 0) {
                let rev = pendientes.shift(); 
                let idx = bloquesDisponibles.shift();
                
                let filaTop = filasWord.items[idx];
                let filaMid = filasWord.items[idx+1];
                let filaBot = filasWord.items[idx+2];

                // Llenamos etiquetas si estaban borradas
                filaTop.cells.items[2].body.insertText("NOMBRE", "Replace");
                filaMid.cells.items[2].body.insertText("FIRMA", "Replace");
                filaBot.cells.items[2].body.insertText("FECHA", "Replace");

                // Datos
                filaTop.cells.items[0].body.insertText(rev.letra, "Replace");
                filaTop.cells.items[1].body.insertText(rev.desc, "Replace");
                
                // Repetimos letra/desc en filas ocultas para consistencia visual si no hay merge
                // (Si ya hay merge visual, esto no se ve, pero no daña)
                filaMid.cells.items[0].body.insertText(rev.letra, "Replace");
                filaBot.cells.items[0].body.insertText(rev.letra, "Replace");

                for(let c=3; c < filaBot.cells.items.length; c++) {
                    filaBot.cells.items[c].body.insertText(rev.fecha, "Replace");
                }
            }

            // D. CREAR NUEVOS BLOQUES
            if (pendientes.length > 0) {
                // Molde
                let nombresMolde = [];
                let anchoTabla = 8;
                if (filasWord.items.length > 0) {
                    let ultimaFila = filasWord.items[filasWord.items.length - 1];
                    ultimaFila.load("values");
                    await context.sync();
                    nombresMolde = ultimaFila.values[0];
                    anchoTabla = nombresMolde.length;
                } else { nombresMolde = new Array(8).fill(""); }

                for (let rev of pendientes) {
                    let f1 = [...nombresMolde]; f1[0]=rev.letra; f1[1]=rev.desc; if(f1.length>2) f1[2]="NOMBRE";
                    let f2 = new Array(anchoTabla).fill(""); f2[0]=rev.letra; f2[1]=rev.desc; if(f2.length>2) f2[2]="FIRMA";
                    let f3 = new Array(anchoTabla).fill(""); f3[0]=rev.letra; f3[1]=rev.desc; if(f3.length>2) f3[2]="FECHA";
                    for(let k=3; k<anchoTabla; k++) { if(nombresMolde[k]!=="") f3[k] = rev.fecha; }

                    let rangoNuevo = tablaWord.addRows("End", 3, [f1, f2, f3]);
                    rangoNuevo.load("rows/cells/body");
                    await context.sync();

                    // Merge Col 0
                    let topRev = rangoNuevo.rows.items[0].cells.items[0].body.getRange("Whole");
                    let botRev = rangoNuevo.rows.items[2].cells.items[0].body.getRange("Whole");
                    topRev.expandTo(botRev).merge();
                    rangoNuevo.rows.items[0].cells.items[0].verticalAlignment = "Center";

                    // Merge Col 1
                    let topDesc = rangoNuevo.rows.items[0].cells.items[1].body.getRange("Whole");
                    let botDesc = rangoNuevo.rows.items[2].cells.items[1].body.getRange("Whole");
                    topDesc.expandTo(botDesc).merge();
                    rangoNuevo.rows.items[0].cells.items[1].verticalAlignment = "Center";
                }
            }
        } 
        
        // =========================================================
        // 🏗️ MODO CODELCO (STACK UP)
        // =========================================================
        else {
            console.log("🟢 MODO ESTÁNDAR ACTIVADO");
            let slotsDisponibles = [];

            for (let i = 0; i < filasWord.items.length; i++) {
                let fila = filasWord.items[i];
                if (fila.cells.items.length < 3) continue;
                
                let texto = fila.cells.items[0].value.trim().toUpperCase();
                
                if (palabrasProtegidas.some(p => texto.includes(p))) continue;

                if (texto === "") {
                    slotsDisponibles.push(fila);
                } 
                else if (mapaDeseado.has(texto)) {
                    let datos = mapaDeseado.get(texto);
                    fila.cells.items[1].body.insertText(datos.fecha, "Replace");
                    fila.cells.items[2].body.insertText(datos.desc, "Replace");
                    mapaDeseado.delete(texto); 
                } 
                else {
                    fila.cells.items[0].body.insertText("", "Replace");
                    fila.cells.items[1].body.insertText("", "Replace");
                    fila.cells.items[2].body.insertText("", "Replace");
                    slotsDisponibles.push(fila);
                }
            }

            let pendientes = [...revisions].reverse().filter(r => mapaDeseado.has(r.letra));
            let nuevasParaCrear = [];

            for (let rev of pendientes) {
                if (slotsDisponibles.length > 0) {
                    let slot = slotsDisponibles.shift();
                    slot.cells.items[0].body.insertText(rev.letra, "Replace");
                    slot.cells.items[1].body.insertText(rev.fecha, "Replace");
                    slot.cells.items[2].body.insertText(rev.desc, "Replace");
                } else {
                    nuevasParaCrear.push([rev.letra, rev.fecha, rev.desc]);
                }
            }

            if (nuevasParaCrear.length > 0) {
                let molde = [];
                if (filasWord.items.length > 0) {
                    filasWord.items[0].load("values");
                    await context.sync();
                    molde = filasWord.items[0].values[0];
                } else { molde = new Array(7).fill(""); }

                const datosInsertar = nuevasParaCrear.map(d => {
                    let f = [...molde];
                    if(f.length>=1) f[0]=d[0];
                    if(f.length>=2) f[1]=d[1];
                    if(f.length>=3) f[2]=d[2];
                    return f;
                });
                tablaWord.addRows("Start", datosInsertar.length, datosInsertar);
            }

            if (slotsDisponibles.length > 0) {
                await context.sync();
                slotsDisponibles.reverse().forEach(f => f.delete());
            }
        }

        await context.sync();
        mostrarMensaje("✅ Tabla Sincronizada.", "green");

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