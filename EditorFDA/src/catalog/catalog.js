/* global Office */

let baseDatosCompleta = [];
let proyectoActual = null;

// URL FIJA DE GITHUB (La que sabemos que funciona)
// Asegúrate de que este archivo existe en tu repo: src/data/proyectos.json
const urlFuenteDatos = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/proyectos.json";

Office.onReady(async () => {
    await cargarDatosIniciales();
    
    // Eventos de los selectores
    document.getElementById("ddlClientes").onchange = filtrarProyectos;
    document.getElementById("ddlProyectos").onchange = seleccionarProyecto;
    
    // Evento del botón buscar (por si lo usas como fallback)
    const btnSearch = document.getElementById("btnSearch");
    if(btnSearch) btnSearch.onclick = buscar;
});

async function cargarDatosIniciales() {
    // URL de tu archivo en GitHub
    const url = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/proyectos.json";

    try {
        console.log("Intentando cargar datos desde:", url);
        
        // Agregamos "?t=" + tiempo para "engañar" al navegador y que no use la caché vieja
        const response = await fetch(url + "?t=" + new Date().getTime());

        if (!response.ok) {
            throw new Error("Error HTTP " + response.status);
        }

        const data = await response.json();
        console.log("Datos recibidos:", data);

        // --- INICIO DE LA CORRECCIÓN ---
        // Aquí decidimos qué variable usar dependiendo de la estructura
        let listaParaUsar = [];

        if (data.body && Array.isArray(data.body)) {
            // Caso 1: Viene desde Power Automate envuelto en "body"
            listaParaUsar = data.body;
        } else if (Array.isArray(data)) {
            // Caso 2: Viene como array directo (formato antiguo o manual)
            listaParaUsar = data;
        } else {
            console.error("Formato JSON no reconocido:", data);
            listaParaUsar = []; // Evitamos que explote el código
        }
        // --- FIN DE LA CORRECCIÓN ---

        // Ahora usamos 'listaParaUsar' que garantizamos que es un Array
        const proyectosFormateados = listaParaUsar.map(item => {
            return {
                // El operador || permite leer el dato aunque cambie mayúsculas/minúsculas
                id: item.id || item.Title || item.ID, 
                nombre: item.nombre || item.NombreProyecto, 
                cliente: item.cliente || item.Cliente,
                division: item.division || item.Division,
                contrato: item.contrato || item.Contrato
            };
        });

        console.log("Proyectos listos:", proyectosFormateados);

        // AQUÍ CONECTAS CON TU UI (Dropdowns, Tablas, etc.)
        // Si tienes una función para llenar el HTML, llámala aquí pasándole 'proyectosFormateados'
        // Ejemplo: actualizarDropdown(proyectosFormateados);

        return proyectosFormateados;

    } catch (error) {
        console.error("Error crítico cargando datos:", error);
    }
}

function filtrarProyectos() {
    const clienteSeleccionado = document.getElementById("ddlClientes").value;
    const ddlProyectos = document.getElementById("ddlProyectos");
    
    ddlProyectos.innerHTML = '<option value="">-- Seleccione Proyecto --</option>';
    document.getElementById("seccionPlantillas").classList.add("oculto");
    document.getElementById("infoProyecto").classList.add("oculto");

    if (!clienteSeleccionado) {
        ddlProyectos.disabled = true;
        return;
    }

    // Filtramos
    const proyectosFiltrados = baseDatosCompleta.filter(p => p.cliente === clienteSeleccionado);

    proyectosFiltrados.forEach(p => {
        let opt = document.createElement("option");
        opt.value = p.id;
        opt.textContent = p.id + " - " + p.nombre;
        ddlProyectos.appendChild(opt);
    });

    ddlProyectos.disabled = false;
}

function seleccionarProyecto() {
    const idProyecto = document.getElementById("ddlProyectos").value;
    
    if (!idProyecto) {
        document.getElementById("seccionPlantillas").classList.add("oculto");
        return;
    }

    proyectoActual = baseDatosCompleta.find(p => p.id === idProyecto);

    if (proyectoActual) {
        document.getElementById("lblNombre").textContent = proyectoActual.nombre;
        document.getElementById("lblAPI").textContent = "API: " + (proyectoActual.api || "N/A");
        document.getElementById("infoProyecto").classList.remove("oculto");
        document.getElementById("seccionPlantillas").classList.remove("oculto");
    }
}

// BÚSQUEDA MANUAL (Por si escriben el ID en el input de texto antiguo)
function buscar() {
    const val = document.getElementById("inputSearch").value;
    const found = baseDatosCompleta.find(p => p.id === val);
    
    if(found) {
        // Si lo encuentran por ID manual, simulamos la selección en los dropdowns
        document.getElementById("ddlClientes").value = found.cliente;
        filtrarProyectos(); // Actualiza la segunda lista
        document.getElementById("ddlProyectos").value = found.id;
        seleccionarProyecto(); // Muestra la info
    } else {
        alert("Proyecto no encontrado en la base de datos.");
    }
}

window.seleccionarPlantilla = function(tipo) {
    if(!proyectoActual) return;
    const mensaje = {
        accion: "CREAR_DOCUMENTO",
        plantilla: tipo,
        datos: proyectoActual
    };
    Office.context.ui.messageParent(JSON.stringify(mensaje));
}