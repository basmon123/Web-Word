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

    // 0. CONFIGURACIÓN
    const estandar = document.getElementById("ddlEstandar").value;
    const esAMSA = (estandar === "AMSA"); // TRUE = Hacia Abajo (3 filas). FALSE = Hacia Arriba (1 fila).

    await Word.run(async (context) => {
        // 1. OBTENER TABLA
        const contentControls = context.document.contentControls.getByTag("ccTablaRevisiones");
        contentControls.load("items/tables/items");
        await context.sync();

        if (contentControls.items.length === 0) {
            mostrarMensaje("❌ No encontré la tabla 'ccTablaRevisiones'.", "red");
            return;
        }

        const tablaWord = contentControls.items[0].tables.items[0];
        const filasWord = tablaWord.rows;
        filasWord.load("items/cells/items/value");
        await context.sync();

        // 2. DETECTAR QUÉ REVISIONES YA EXISTEN (Para no duplicar)
        let letrasEnTabla = new Set();
        for (let i = 0; i < filasWord.items.length; i++) {
            // Leemos la primera celda de cada fila. 
            // .trim() evita errores por espacios vacíos
            let val = filasWord.items[i].cells.items[0].value.trim().toUpperCase();
            // Ignoramos encabezados o filas vacías
            if (val && val !== "REV" && val !== "REVISIÓN" && val !== "REVISION") {
                letrasEnTabla.add(val);
            }
        }

        // 3. FILTRAR: ¿QUÉ FALTA POR ESCRIBIR?
        // Solo procesamos las que están en tu lista pero NO en Word.
        let nuevasParaEscribir = revisions.filter(r => !letrasEnTabla.has(r.letra));

        if (nuevasParaEscribir.length === 0) {
            mostrarMensaje("✅ La tabla ya está actualizada.", "green");
            return;
        }

        console.log(`Modo AMSA: ${esAMSA}. Insertando ${nuevasParaEscribir.length} nuevas revisiones.`);

        // =========================================================
        // 🏗️ CASO 1: MODO AMSA (HACIA ABAJO - STACK DOWN - 3 FILAS)
        // =========================================================
        if (esAMSA) {
            
            // A. OBTENER MOLDE DE NOMBRES (De la última fila existente)
            let nombresMolde = [];
            let anchoTabla = 8; // Valor seguro por defecto

            if (filasWord.items.length > 0) {
                // En Stack Down, la última revisión está al final.
                // Si la tabla tiene filas, copiamos la estructura de la última.
                // Si es un bloque de 3, la fila de nombres suele ser la primera del bloque (antepenúltima total)
                // Pero para ir a lo seguro, tomamos la ÚLTIMA fila visible y copiamos sus firmas.
                let ultimaFila = filasWord.items[filasWord.items.length - 1];
                ultimaFila.load("values");
                await context.sync();
                
                // values[0] es el array de datos ["Letra", "Desc", "Etiqueta", "Firma1", "Firma2"...]
                nombresMolde = ultimaFila.values[0]; 
                anchoTabla = nombresMolde.length;
            } else {
                nombresMolde = new Array(8).fill("");
            }

            // B. ITERAR Y CREAR UNO POR UNO (Al final de la tabla)
            for (let rev of nuevasParaEscribir) {
                
                // Construimos las 3 filas del bloque AMSA
                
                // Fila 1: NOMBRES (Copiamos firmas del molde)
                let fila1 = [...nombresMolde]; 
                fila1[0] = rev.letra; 
                fila1[1] = rev.desc; 
                if(fila1.length > 2) fila1[2] = "NOMBRE"; // Etiqueta
                // El resto (indices 3,4,5...) se mantienen con los nombres clonados (Juan, Pedro...)

                // Fila 2: FIRMAS (Vacío para firmar a mano)
                let fila2 = new Array(anchoTabla).fill("");
                fila2[0] = rev.letra; 
                fila2[1] = rev.desc; 
                if(fila2.length > 2) fila2[2] = "FIRMA";

                // Fila 3: FECHAS
                let fila3 = new Array(anchoTabla).fill("");
                fila3[0] = rev.letra; 
                fila3[1] = rev.desc; 
                if(fila3.length > 2) fila3[2] = "FECHA";
                // Rellenamos la fecha en las columnas de firmas (asumimos desde la col 3)
                for(let k=3; k<anchoTabla; k++) { 
                    // Ponemos fecha solo si esa columna tenía nombre en el molde (para no ensuciar celdas vacías)
                    if (nombresMolde[k] !== "") fila3[k] = rev.fecha; 
                }

                // C. INSERTAR "End" (AL FINAL)
                // addRows devuelve un Rango con las filas nuevas
                let rangoNuevo = tablaWord.addRows("End", 3, [fila1, fila2, fila3]);
                
                // Sincronizamos para que el rango exista físicamente
                await context.sync();

                // D. FUSIONAR CELDAS (MERGE) 🧬
                // Usamos getCell(fila, columna) relativo al nuevo rango (indices 0, 1, 2)
                
                // Columna 0 (REV): Fusionar fila 0 con fila 2 del nuevo rango
                let celdaRevTop = rangoNuevo.getCell(0, 0);
                let celdaRevBot = rangoNuevo.getCell(2, 0);
                celdaRevTop.merge(celdaRevBot);
                celdaRevTop.verticalAlignment = "Center"; // Centrar verticalmente

                // Columna 1 (DESC): Fusionar fila 0 con fila 2 del nuevo rango
                let celdaDescTop = rangoNuevo.getCell(0, 1);
                let celdaDescBot = rangoNuevo.getCell(2, 1);
                celdaDescTop.merge(celdaDescBot);
                celdaDescTop.verticalAlignment = "Center";
            }
        } 
        
        // =========================================================
        // 🏗️ CASO 2: MODO CODELCO/ESTÁNDAR (HACIA ARRIBA - 1 FILA)
        // =========================================================
        else {
            
            // A. OBTENER MOLDE DE NOMBRES (De la PRIMERA fila existente)
            let nombresMolde = [];
            
            if (filasWord.items.length > 0) {
                // En Stack Up, la revisión más reciente está arriba (índice 0).
                // Copiamos sus firmas para la nueva que irá ENCIMA de ella.
                let primeraFila = filasWord.items[0];
                primeraFila.load("values");
                await context.sync();
                nombresMolde = primeraFila.values[0];
            } else {
                nombresMolde = new Array(7).fill("");
            }

            // B. CONSTRUIR DATOS (Invertimos orden para insertarlas bien si hay varias)
            // Si hay que insertar C y D, y usamos "Start", primero insertamos C, luego D encima.
            // Así queda D, C, B, A.
            // 'nuevasParaEscribir' viene en orden A,B,C... lo invertimos si queremos stack up estricto, 
            // o lo iteramos normal. Para "Start", iteramos normal:
            // 1. Inserto C en 0. (Tabla: C, B, A)
            // 2. Inserto D en 0. (Tabla: D, C, B, A) -> Correcto.
            
            // Usaremos map para preparar todas las filas de golpe
            const datosParaInsertar = nuevasParaEscribir.map(rev => {
                let filaNueva = [...nombresMolde]; // Copia del molde
                
                // Sobrescribimos datos nuevos
                if(filaNueva.length >= 1) filaNueva[0] = rev.letra;
                if(filaNueva.length >= 2) filaNueva[1] = rev.fecha;
                if(filaNueva.length >= 3) filaNueva[2] = rev.desc;
                
                return filaNueva;
            });

            // C. INSERTAR "Start" (AL PRINCIPIO)
            // Insertamos todas de una vez, pero ojo con el orden visual.
            // Si mandamos el bloque [C, D], "Start" pondrá [C, D] y empujará lo viejo.
            // Quedará: C, D, B, A.
            // Si queremos D, C, B, A, tenemos que invertir el array de datos antes de enviarlo.
            datosParaInsertar.reverse(); 

            if (datosParaInsertar.length > 0) {
                tablaWord.addRows("Start", datosParaInsertar.length, datosParaInsertar);
            }
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