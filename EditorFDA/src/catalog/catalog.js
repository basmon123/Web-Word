/* global Office */

let baseDatosCompleta = [];
let proyectoActual = null;

// ==========================================
// 1. CONFIGURACIÓN
// ==========================================
// 🔴 RECUERDA PONER AQUÍ TU URL DE SHAREPOINT
const urlFuenteDatos = "PON_AQUI_LA_URL_DE_TU_JSON_EN_SHAREPOINT"; 

// DICCIONARIO DE REGLAS
// El código buscará estas palabras clave. Si las encuentra, asigna el Cliente Global.
const CONFIG_CLIENTES = {
    "CODELCO CHILE": [
        "Codelco", "Chuquicamata", "Gabriela Mistral", "Ministro Hales", 
        "Teniente", "Ventanas", "Vicepresidencia", "Distrito Norte", "Casa Matriz"
    ],
    "AMSA": [
        "Centinela", "Pelambres", "Antofagasta Minerals", "AMSA"
    ],
    "AGUAS ANDINAS": ["Aguas Andina"],
    "COMPAÑÍA INDUSTRIAL VOLCÁN S.A.": ["Volcan", "Volcán"],
    "COMPLEJO METALÚRGICO ALTONORTE": ["Altonorte"],
    "CONSULTING SKAVA": ["SKAVA"],
    "CORPORACIÓN ABB": ["ABB"],
    "ENAMI": ["ENAMI", "Empresa Nacional de Minería"],
    "ESMAX": ["Esmax"],
    "GERENCIA DE PROYECTOS DISTRITAL": ["Proyectos Distrital"],
    "GESVIAL": ["Gesvial", "Gestión Vial"],
    "INPPA": ["INPPA"],
    "KAPSCH TRAFFICCOM": ["Kapsch"],
    "PATERSON & COOKE": ["Paterson"],
    "ROBERT BOSCH CHILE S.A.": ["Bosch"],
    "SOCIEDAD PUNTA DEL COBRE S.A.": ["Punta del Cobre"],
    "STATKRAFT": ["Statkraft"],
    "TRANSELEC": ["Transelec"]
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

// ==========================================
// 3. LÓGICA INTELIGENTE (PROCESAMIENTO)
// ==========================================

function procesarCliente(nombreRaw) {
    if (!nombreRaw) return { global: "OTROS", division: "---" };

    const nombreMayus = nombreRaw.toUpperCase();

    // 1. Buscar coincidencia en el diccionario
    for (const [clienteGlobal, palabrasClave] of Object.entries(CONFIG_CLIENTES)) {
        const encontrado = palabrasClave.some(palabra => nombreMayus.includes(palabra.toUpperCase()));
        
        if (encontrado) {
            // ¡ENCONTRADO! Ahora limpiamos el nombre de la división
            let divisionLimpia = nombreRaw;

            // Lógica especial para CODELCO: Quitamos la palabra "Codelco" y "División" del nombre
            if (clienteGlobal === "CODELCO CHILE") {
                // El regex /Codelco/gi busca la palabra ignorando mayúsculas/minúsculas y la borra
                divisionLimpia = divisionLimpia.replace(/Codelco/gi, "").replace(/División/gi, "").trim();
                
                // Si al limpiar queda vacío (ej: el input era solo "Codelco"), ponemos algo genérico
                if (divisionLimpia === "" || divisionLimpia === "-") divisionLimpia = "General";
                
                // Capitalizar primera letra (Estética: "casa matriz" -> "Casa Matriz")
                divisionLimpia = divisionLimpia.charAt(0).toUpperCase() + divisionLimpia.slice(1);
            }

            return { 
                global: clienteGlobal, 
                division: divisionLimpia 
            };
        }
    }

    // 2. Si no está en la lista, se va a "OTROS" o usa su propio nombre
    return { global: nombreRaw, division: "---" };
}

async function cargarDatosIniciales() {
    try {
        const response = await fetch(urlFuenteDatos + "?t=" + new Date().getTime(), {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });

        if (!response.ok) throw new Error(`Error HTTP: ${response.status}`);

        const data = await response.json();
        
        let lista = [];
        if (Array.isArray(data)) lista = data;
        else if (data.body && Array.isArray(data.body)) lista = data.body;
        else if (data.value && Array.isArray(data.value)) lista = data.value;

        // MAPEO
        baseDatosCompleta = lista.map(item => {
            const rawCliente = item.cliente || "";
            // Aquí ocurre la magia de la limpieza
            const infoCliente = procesarCliente(rawCliente);

            return {
                id: item.id,
                nombre: item.nombre,
                cliente: infoCliente.global,     // Ej: CODELCO CHILE
                division: infoCliente.division,  // Ej: Casa Matriz (sin "Codelco" antes)
                contrato: item.contrato,
                carpeta_plantilla: item.carpeta_plantilla,
                api: item.api || ""
            };
        });

        // Llenar Dropdown Clientes
        const ddlClientes = document.getElementById("ddlClientes");
        ddlClientes.innerHTML = '<option value="">-- Seleccione Cliente --</option>';
        
        const clientesUnicos = [...new Set(baseDatosCompleta.map(p => p.cliente))].sort();
        
        clientesUnicos.forEach(c => {
            if(c && c !== "OTROS") { // Opcional: Ocultar "Otros" o dejarlo al final
                let opt = document.createElement("option");
                opt.value = c;
                opt.textContent = c;
                ddlClientes.appendChild(opt);
            }
        });

    } catch (error) {
        console.error("Error cargando datos:", error);
        document.getElementById("ddlClientes").innerHTML = '<option>Error al cargar datos</option>';
    }
}

// ==========================================
// 4. LÓGICA DE INTERFAZ (IGUAL QUE ANTES)
// ==========================================

function filtrarProyectos() {
    const clienteSel = document.getElementById("ddlClientes").value;
    const ddlProyectos = document.getElementById("ddlProyectos");
    
    ddlProyectos.innerHTML = '<option value="">-- Seleccione N° --</option>';
    ocultarDetalles();

    if (!clienteSel) {
        ddlProyectos.disabled = true;
        return;
    }

    const filtrados = baseDatosCompleta.filter(p => p.cliente === clienteSel);

    // Ordenar los proyectos numéricamente si son números
    filtrados.sort((a, b) => a.id - b.id);

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
        setText("lblDivision", proyectoActual.division); // Ahora saldrá limpio
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

window.seleccionarPlantilla = function(tipo) {
    if(!proyectoActual) return;
    localStorage.setItem("FDA_ProyectoActual", JSON.stringify(proyectoActual));
    const mensaje = {
        accion: "CREAR_DOCUMENTO",
        plantilla: tipo,
        datos: proyectoActual
    };
    if (Office.context.ui.messageParent) {
        Office.context.ui.messageParent(JSON.stringify(mensaje));
    }
}