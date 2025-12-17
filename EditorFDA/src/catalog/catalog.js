/* global Office */

let baseDatosCompleta = [];
let listaPlantillas = []; // <--- NUEVA VARIABLE PARA PLANTILLAS
let proyectoActual = null;

// ==========================================
// 1. CONFIGURACIÓN
// ==========================================

// URL DE PROYECTOS (GitHub)
const urlFuenteDatos = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/proyectos.json";

// 🔴 URL DE PLANTILLAS (GitHub) - ¡ACTUALIZA ESTO!
const urlFuentePlantillas = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/plantillas.json";

// DICCIONARIO DE REGLAS
const CONFIG_CLIENTES = {
    "CODELCO CHILE": [
        "Codelco", "Chuquicamata", "Gabriela Mistral", "Ministro Hales", 
        "Teniente", "Ventanas", "Vicepresidencia", "Distrito Norte", "Casa Matriz", 
        "Salvador", "Radomiro Tomic", "Andina", "Proyectos Distrital"
    ],
    "AMSA": [
        "Centinela", "Pelambres", "Antofagasta Minerals", "AMSA", "Antucoya", "Zaldivar"
    ],
    "AGUAS ANDINAS": ["Aguas Andina"],
    "COMPAÑÍA INDUSTRIAL VOLCÁN S.A.": ["Volcan", "Volcán"],
    "COMPLEJO METALÚRGICO ALTONORTE": ["Altonorte"],
    "CONSULTING SKAVA": ["SKAVA"],
    "CORPORACIÓN ABB": ["ABB"],
    "ENAMI": ["ENAMI", "Empresa Nacional de Minería"],
    "ESMAX": ["Esmax"],
    "GESVIAL": ["Gesvial", "Gestión Vial"],
    "INPPA": ["INPPA"],
    "KAPSCH TRAFFICCOM": ["Kapsch"],
    "PATERSON & COOKE": ["Paterson"],
    "ROBERT BOSCH CHILE S.A.": ["Bosch"],
    "SOCIEDAD PUNTA DEL COBRE S.A.": ["Punta del Cobre"],
    "STATKRAFT": ["Statkraft"],
    "TRANSELEC": ["Transelec"],
    "BHP BILLITON": ["Escondida", "Spence", "Cerro Colorado", "BHP"]
};

// ==========================================
// 2. INICIO
// ==========================================

Office.onReady(async () => {
    await cargarDatosIniciales();
    
    const ddlClientes = document.getElementById("ddlClientes");
    const ddlProyectos = document.getElementById("ddlProyectos");

    if(ddlClientes) ddlClientes.onchange = filtrarProyectos;
    if(ddlProyectos) ddlProyectos.onchange = seleccionarProyecto;
});

/// ==========================================
// 3. LÓGICA DE DATOS
// ==========================================

function procesarCliente(nombreRaw) {
    if (!nombreRaw) return { global: "OTROS", division: "---" };

    const nombreMayus = nombreRaw.toUpperCase();

    for (const [clienteGlobal, palabrasClave] of Object.entries(CONFIG_CLIENTES)) {
        const encontrado = palabrasClave.some(palabra => nombreMayus.includes(palabra.toUpperCase()));
        
        if (encontrado) {
            let divisionLimpia = nombreRaw;

            if (clienteGlobal === "CODELCO CHILE") {
                divisionLimpia = divisionLimpia.replace(/Codelco/gi, "").replace(/División/gi, "").trim();
                if (divisionLimpia === "" || divisionLimpia === "-") divisionLimpia = "General";
                if (divisionLimpia.startsWith("-")) divisionLimpia = divisionLimpia.substring(1).trim();
            }

            return { 
                global: clienteGlobal, 
                division: divisionLimpia.toUpperCase()
            };
        }
    }
    return { global: nombreRaw.toUpperCase(), division: "---" };
}

async function cargarDatosIniciales() {
    try {
        const timestamp = new Date().getTime();

        // 1. Cargar Proyectos
        const resProyectos = await fetch(urlFuenteDatos + "?t=" + timestamp);
        const dataProyectos = await resProyectos.json();

        // 2. Cargar Plantillas (CON PROTECCIÓN CONTRA EL ERROR QUE TE SALIÓ)
        try {
            const resPlantillas = await fetch(urlFuentePlantillas + "?t=" + timestamp);
            if(resPlantillas.ok) {
                const rawPlantillas = await resPlantillas.json();
                
                // 🛡️ AQUÍ ESTÁ EL ARREGLO:
                // Si viene como lista [...], lo usamos.
                if (Array.isArray(rawPlantillas)) {
                    listaPlantillas = rawPlantillas;
                } 
                // Si viene como objeto único {...}, lo metemos en una lista.
                else if (typeof rawPlantillas === 'object' && rawPlantillas !== null) {
                    listaPlantillas = [rawPlantillas];
                }
                
                console.log("Plantillas cargadas:", listaPlantillas.length);
            }
        } catch(e) {
            console.warn("No se pudo cargar plantillas.json", e);
        }
        
        // Normalizar estructura proyectos
        let lista = [];
        if (Array.isArray(dataProyectos)) lista = dataProyectos;
        else if (dataProyectos.body && Array.isArray(dataProyectos.body)) lista = dataProyectos.body;
        else if (dataProyectos.value && Array.isArray(dataProyectos.value)) lista = dataProyectos.value;

        // Mapeo
        baseDatosCompleta = lista.map(item => {
            const rawCliente = item.cliente || "";
            const infoCliente = procesarCliente(rawCliente);

            return {
                id: item.id,
                nombre: (item.nombre || "").toUpperCase(),
                cliente: infoCliente.global,     
                division: infoCliente.division, 
                contrato: (item.contrato || "").toUpperCase(),
                carpeta_plantilla: item.carpeta_plantilla, 
                api: item.api || ""
            };
        });

        // Llenar Dropdown
        const ddlClientes = document.getElementById("ddlClientes");
        ddlClientes.innerHTML = '<option value="">-- SELECCIONE CLIENTE --</option>';
        
        const clientesUnicos = [...new Set(baseDatosCompleta.map(p => p.cliente))].sort();
        
        clientesUnicos.forEach(c => {
            if(c && c !== "OTROS") { 
                let opt = document.createElement("option");
                opt.value = c;
                opt.textContent = c;
                ddlClientes.appendChild(opt);
            }
        });

    } catch (error) {
        console.error("Error cargando datos:", error);
        document.getElementById("ddlClientes").innerHTML = '<option>ERROR AL CARGAR DATOS</option>';
    }
}

// ==========================================
// 4. EL BUSCADOR INTELIGENTE (VERSIÓN ROBUSTA)
// ==========================================

// Función auxiliar para estandarizar textos antes de comparar
function normalizar(texto) {
    if (!texto) return ""; // Si es null o undefined, devuelve vacío
    const t = String(texto).trim().toUpperCase(); // Todo a mayúsculas y sin espacios extra
    if (t === "-" || t === "---" || t === "null") return ""; // Tratar guiones como vacío
    return t;
}

function buscarUrlPlantilla(tipoDocumento, proyecto) {
    if (!listaPlantillas || listaPlantillas.length === 0) {
        console.error("❌ Error: La lista de plantillas está vacía o no cargó.");
        return null;
    }

    // 1. Preparamos los datos del PROYECTO (Lo que buscamos)
    const tipoReq = normalizar(tipoDocumento);
    const pContrato = normalizar(proyecto.contrato);
    const pDivision = normalizar(proyecto.division);
    const pCliente = normalizar(proyecto.cliente);

    console.log(`🔍 BUSCANDO PLANTILLA: [${tipoReq}]`);
    console.log(`   Datos Proyecto -> Cliente: ${pCliente} | División: ${pDivision} | Contrato: ${pContrato}`);

    let encontrada = null;

    // ---------------------------------------------------------
    // 🥇 PRIORIDAD 1: CONTRATO (Solo si el proyecto tiene contrato)
    // ---------------------------------------------------------
    if (pContrato !== "") {
        encontrada = listaPlantillas.find(p => {
            const tTipo = normalizar(p.tipo);
            const tContrato = normalizar(p.contrato);
            // Coincide Tipo Y Coincide Contrato
            return tTipo === tipoReq && tContrato === pContrato;
        });

        if (encontrada) console.log("✅ Encontrada por CONTRATO:", encontrada.nombre);
    }

    // ---------------------------------------------------------
    // 🥈 PRIORIDAD 2: DIVISIÓN (Solo si el proyecto tiene división)
    // ---------------------------------------------------------
    if (!encontrada && pDivision !== "") {
        encontrada = listaPlantillas.find(p => {
            const tTipo = normalizar(p.tipo);
            const tDivision = normalizar(p.division);
            // Coincide Tipo Y Coincide División
            return tTipo === tipoReq && tDivision === pDivision;
        });

        if (encontrada) console.log("✅ Encontrada por DIVISIÓN:", encontrada.nombre);
    }

    // ---------------------------------------------------------
    // 🥉 PRIORIDAD 3: CLIENTE GLOBAL
    // (Buscamos que coincida el cliente y que la plantilla NO sea específica de otra cosa)
    // ---------------------------------------------------------
    if (!encontrada) {
        encontrada = listaPlantillas.find(p => {
            const tTipo = normalizar(p.tipo);
            const tCliente = normalizar(p.cliente);
            const tDivision = normalizar(p.division);
            const tContrato = normalizar(p.contrato);

            // Coincide Tipo Y Coincide Cliente Y (Division vacía) Y (Contrato vacío)
            return tTipo === tipoReq && 
                   tCliente === pCliente && 
                   tDivision === "" && 
                   tContrato === "";
        });

        if (encontrada) console.log("✅ Encontrada por CLIENTE GLOBAL:", encontrada.nombre);
    }

    // ---------------------------------------------------------
    // 🛡️ PRIORIDAD 4: GENERAL (Fallback)
    // ---------------------------------------------------------
    if (!encontrada) {
        encontrada = listaPlantillas.find(p => {
            const tTipo = normalizar(p.tipo);
            const tCliente = normalizar(p.cliente);
            
            return tTipo === tipoReq && tCliente === "GENERAL";
        });

        if (encontrada) console.log("✅ Encontrada por GENERAL:", encontrada.nombre);
    }

    // Resultado final
    if (!encontrada) {
        console.warn("⚠️ NO SE ENCONTRÓ NINGUNA PLANTILLA COMPATIBLE.");
        console.log("   --> Revisa tu archivo plantillas.json en GitHub.");
        console.log("   --> Asegúrate de que las columnas coincidan.");
    }

    return encontrada ? encontrada.url : null;
}
// ==========================================
// 5. LÓGICA DE INTERFAZ
// ==========================================

function filtrarProyectos() {
    const clienteSel = document.getElementById("ddlClientes").value;
    const ddlProyectos = document.getElementById("ddlProyectos");
    
    ddlProyectos.innerHTML = '<option value="">-- SELECCIONE N° --</option>';
    ocultarDetalles();

    if (!clienteSel) {
        ddlProyectos.disabled = true;
        return;
    }

    const filtrados = baseDatosCompleta.filter(p => p.cliente === clienteSel);

    filtrados.sort((a, b) => {
        const numA = parseInt(a.id, 10);
        const numB = parseInt(b.id, 10);
        if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
        return a.id.localeCompare(b.id);
    });

    filtrados.forEach(p => {
        let opt = document.createElement("option");
        opt.text = p.id; 
        opt.value = p.id;
        ddlProyectos.appendChild(opt);
    });

    ddlProyectos.disabled = false;
}

function seleccionarProyecto() {
    const idProyecto = document.getElementById("ddlProyectos").value;
    
    if (!idProyecto) {
        ocultarDetalles();
        return;
    }

    proyectoActual = baseDatosCompleta.find(p => String(p.id) === String(idProyecto));

    if (proyectoActual) {
        setText("lblNombre", proyectoActual.nombre);
        setText("lblCliente", proyectoActual.cliente);
        setText("lblDivision", proyectoActual.division);
        setText("lblContrato", proyectoActual.contrato);
        setText("lblAPI", proyectoActual.api);

        document.getElementById("infoProyecto").classList.remove("oculto");
        document.getElementById("seccionPlantillas").classList.remove("oculto");
    }
}

function ocultarDetalles() {
    document.getElementById("infoProyecto").classList.add("oculto");
    document.getElementById("seccionPlantillas").classList.add("oculto");
    proyectoActual = null;
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text || "---";
}

// 🔴 FUNCIÓN ACTUALIZADA: ENVÍA URL EN LUGAR DE TIPO
window.seleccionarPlantilla = function(tipo) {
    if(!proyectoActual) return;

    // 1. Buscamos la URL inteligente
    const urlFinal = buscarUrlPlantilla(tipo, proyectoActual);

    if (!urlFinal) {
        // Fallback visual si no hay plantilla
        console.error("No se encontró plantilla para", tipo);
        // Opcional: Mostrar alerta al usuario
        return; 
    }

    localStorage.setItem("FDA_ProyectoActual", JSON.stringify(proyectoActual));
    
    const mensaje = {
        accion: "CREAR_DOCUMENTO",
        plantillaUrl: urlFinal, // <--- Enviamos la URL directa
        datos: proyectoActual
    };
    
    if (Office.context.ui.messageParent) {
        Office.context.ui.messageParent(JSON.stringify(mensaje));
    }
}