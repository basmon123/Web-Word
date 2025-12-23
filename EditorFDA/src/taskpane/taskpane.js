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

/// ---------------------------------------------
// 3. ESCRITURA EN WORD (INTELIGENTE: ACTUALIZA O INSERTA)
// ---------------------------------------------
async function escribirTablaEnWord() {
    mostrarMensaje("⏳ Sincronizando tabla...", "blue");

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
        
        // Carga inicial
        let filasWord = tablaWord.rows;
        filasWord.load("items/cells/items/value, items/cells/items/body, items/values");
        await context.sync();

        // =========================================================
        // 🕵️‍♂️ DETECTOR SOLO POR DIRECCIÓN (STACK UP vs DOWN)
        // =========================================================
        
        // Empezamos respetando lo que eligió el usuario
        let estandarSeleccionado = document.getElementById("ddlEstandar").value;
        let esAMSA = (estandarSeleccionado === "AMSA");
        
        // Buscamos índices de A y B para saber la dirección
        let indexA = -1;
        let indexB = -1;

        for (let i = 0; i < filasWord.items.length; i++) {
            let row = filasWord.items[i];
            if (!row.cells || !row.cells.items || row.cells.items.length === 0) continue;
            
            // Miramos la primera celda
            let val = row.cells.items[0].value.trim().toUpperCase();
            if (val === "A" || val === "0") indexA = i;
            if (val === "B" || val === "1") indexB = i;
        }

        // TOMA DE DECISIÓN (Solo si hay evidencia clara de dirección)
        if (indexA !== -1 && indexB !== -1) {
            if (indexA < indexB) {
                // A está ARRIBA de B -> STACK DOWN (Crecen hacia abajo) -> MODO AMSA
                if (!esAMSA) {
                    console.log("🤖 AUTO: Detectado Stack Down (A sobre B). Cambiando a AMSA.");
                    esAMSA = true;
                    document.getElementById("ddlEstandar").value = "AMSA";
                }
            } else {
                // A está ABAJO de B -> STACK UP (Crecen hacia arriba) -> MODO CODELCO
                if (esAMSA) {
                    console.log("🤖 AUTO: Detectado Stack Up (A bajo B). Cambiando a CODELCO.");
                    esAMSA = false;
                    document.getElementById("ddlEstandar").value = "CODELCO";
                }
            }
        }
        // SI NO HAY A y B JUNTAS: NO HACEMOS NADA. CONFIAMOS EN EL USUARIO.

        const palabrasProtegidas = [
            "REVISIÓN", "REVISION", "REV", "REV.", "FECHA", "EMITIDO", "DESCRIPCIÓN", "DESCRIPCION", 
            "PROYECTO", "TÍTULO", "TITULO", "FDA", "CENTINELA", "APROBÓ", "APROBO", "PREPARÓ", 
            "PREPARO", "POR", "REVISÓ", "REVISO", "CLIENTE", "N°", "N.", "NO.", "INTERNO", "NOMBRE", "FIRMA"
        ];

        // =========================================================
        // 🏗️ MODO AMSA (STACK DOWN - CRECE HACIA ABAJO)
        // =========================================================
        if (esAMSA) {
            console.log("🔵 MODO AMSA (Stack Down): ACTIVADO");

            // 1. MOLDE DE NOMBRES
            let moldeNombres = new Array(8).fill(""); 
            for (let i = filasWord.items.length - 1; i >= 0; i--) {
                let row = filasWord.items[i];
                if (!row || !row.cells || row.cells.items.length < 4) continue;
                
                let etiquetaLateral = "";
                let valorNombre = "";
                try {
                    if(row.cells.items[2]) etiquetaLateral = row.cells.items[2].value.trim().toUpperCase();
                    if(row.cells.items[3]) valorNombre = row.cells.items[3].value.trim();
                } catch(e) { continue; }

                // Ancla "NOMBRE" estricta
                if (etiquetaLateral.includes("NOMB") && valorNombre.length > 0 && 
                    !valorNombre.includes("<") && !valorNombre.includes("/") &&
                    !palabrasProtegidas.some(p => valorNombre.toUpperCase() === p)) {
                    moldeNombres = row.values[0];
                    break; 
                }
            }

            // 2. IDENTIFICAR SLOTS
            let slotsIndices = [];
            for (let i = 0; i < filasWord.items.length; i++) {
                if (!filasWord.items[i].cells || filasWord.items[i].cells.items.length === 0) continue;
                let texto = filasWord.items[i].cells.items[0].value.trim().toUpperCase();
                if (palabrasProtegidas.some(p => texto.includes(p))) continue;
                if (i + 2 < filasWord.items.length) { slotsIndices.push(i); i += 2; }
            }

            // 3. RECICLAR
            let revisionIndex = 0;
            while (revisionIndex < revisions.length && revisionIndex < slotsIndices.length) {
                let rev = revisions[revisionIndex];
                let idx = slotsIndices[revisionIndex]; 
                let filaTop = filasWord.items[idx];     
                let filaMid = filasWord.items[idx+1];   
                let filaBot = filasWord.items[idx+2];   

                // Llenar Datos
                try { 
                    if(filaTop.cells.items.length > 0) filaTop.cells.items[0].body.insertText(rev.letra, "Replace"); 
                    if(filaTop.cells.items.length > 1) filaTop.cells.items[1].body.insertText(rev.desc, "Replace"); 
                    if(filaTop.cells.items.length > 2) filaTop.cells.items[2].body.insertText("NOMBRE", "Replace");
                    for(let c = 3; c < filaTop.cells.items.length; c++) {
                        let esColumnaCliente = (c >= filaTop.cells.items.length - 2);
                        if (rev.letra === "A" && esColumnaCliente) filaTop.cells.items[c].body.insertText("", "Replace");
                        else if (moldeNombres[c] && moldeNombres[c].trim() !== "") filaTop.cells.items[c].body.insertText(moldeNombres[c], "Replace");
                    }
                } catch(e){}
                
                // Etiquetas Firma/Fecha
                try { if(filaMid.cells.items.length > 0) {
                        let idx = (filaMid.cells.items.length < 7) ? 0 : 2;
                        if(filaMid.cells.items.length > idx) filaMid.cells.items[idx].body.insertText("FIRMA", "Replace"); 
                }} catch(e){}
                try { if(filaBot.cells.items.length > 0) {
                        let idx = (filaBot.cells.items.length < 7) ? 0 : 2;
                        if(filaBot.cells.items.length > idx) filaBot.cells.items[idx].body.insertText("FECHA", "Replace"); 
                }} catch(e){}

                // Fechas
                try { 
                    let startCol = (filaBot.cells.items.length < 7) ? 1 : 3;
                    for(let c = startCol; c < filaBot.cells.items.length; c++) {
                        let esColumnaCliente = (c >= filaBot.cells.items.length - 2); 
                        if (rev.letra === "A" && esColumnaCliente) filaBot.cells.items[c].body.insertText("", "Replace"); 
                        else filaBot.cells.items[c].body.insertText(rev.fecha, "Replace"); 
                    }
                } catch(e){}
                revisionIndex++;
            }

            // 4. LIMPIAR
            while (revisionIndex < slotsIndices.length) {
                let idx = slotsIndices[revisionIndex];
                [filasWord.items[idx], filasWord.items[idx+1], filasWord.items[idx+2]].forEach(fila => {
                    try { if(fila.cells) fila.cells.items.forEach(c => c.body.insertText("", "Replace")); } catch(e){}
                });
                revisionIndex++;
            }

            // 5. CREAR NUEVOS BLOQUES (SOLUCIÓN CRASH) 🛠️
            let pendientes = revisions.slice(revisionIndex); 

            if (pendientes.length > 0) {
                let anchoTabla = moldeNombres.length > 0 ? moldeNombres.length : 8;

                for (let rev of pendientes) {
                    // A. PREPARAR DATOS
                    let f1 = [...moldeNombres]; 
                    f1[0]=rev.letra; f1[1]=rev.desc; if(f1.length>2) f1[2]="NOMBRE";
                    let f2 = new Array(anchoTabla).fill(""); if(f2.length>2) f2[2]="FIRMA";
                    let f3 = new Array(anchoTabla).fill(""); if(f3.length>2) f3[2]="FECHA";
                    
                    for(let k=3; k<anchoTabla; k++) { 
                        let esCliente = (k >= 6);
                        if (rev.letra === "A" && esCliente) { f1[k] = ""; f3[k] = ""; } 
                        else { f3[k] = rev.fecha; }
                    }

                    // B. AGREGAR FILAS (Sin capturar variable)
                    tablaWord.addRows("End", 3, [f1, f2, f3]);
                    
                    // C. SINCRONIZAR 
                    await context.sync(); 

                    // D. RECARGAR REFERENCIA DESDE LA TABLA PADRE (Esto evita el InvalidArgument)
                    // Pedimos las filas de nuevo, ahora incluirán las recién creadas
                    let filasRefrescadas = tablaWord.rows;
                    filasRefrescadas.load("items/cells/body"); 
                    await context.sync();

                    // E. MERGE (Usando la referencia refrescada)
                    let total = filasRefrescadas.items.length;
                    if (total >= 3) {
                        let rowTop = filasRefrescadas.items[total - 3];
                        let rowBot = filasRefrescadas.items[total - 1];
                        
                        try {
                            // Merge Rev
                            rowTop.cells.items[0].merge(rowBot.cells.items[0]);
                            rowTop.cells.items[0].verticalAlignment = "Center";
                            // Merge Desc
                            rowTop.cells.items[1].merge(rowBot.cells.items[1]);
                            rowTop.cells.items[1].verticalAlignment = "Center";
                        } catch(e) { console.warn("Merge falló:", e); }
                    }
                }
            }
        } 
        
        // =========================================================
        // 🏗️ MODO CODELCO / ESTÁNDAR (STACK UP - CRECE HACIA ARRIBA)
        // =========================================================
        else {
            console.log("🟢 MODO ESTÁNDAR ACTIVADO");
            let slotsDisponibles = [];
            let mapaDeseado = new Map();
            revisions.forEach(r => mapaDeseado.set(r.letra, r));

            // Escanear filas simples
            for (let i = 0; i < filasWord.items.length; i++) {
                let fila = filasWord.items[i];
                if (fila.cells.items.length < 3) continue;
                let texto = fila.cells.items[0].value.trim().toUpperCase();
                if (palabrasProtegidas.some(p => texto.includes(p))) continue;

                if (texto === "") slotsDisponibles.push(fila);
                else if (mapaDeseado.has(texto)) {
                    let datos = mapaDeseado.get(texto);
                    fila.cells.items[1].body.insertText(datos.fecha, "Replace");
                    fila.cells.items[2].body.insertText(datos.desc, "Replace");
                    mapaDeseado.delete(texto); 
                } 
                else {
                    // Limpieza simple
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
                // Molde simple (fila 0)
                let molde = [];
                if (filasWord.items.length > 0) {
                    // Aquí no hay "NOMBRE" ni nada complejo, solo copiamos la primera fila como estilo
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