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
    try {
        console.log("Intentando cargar datos desde:", urlFuenteDatos);
        
        const response = await fetch(urlFuenteDatos);
        
        if (!response.ok) {
            throw new Error(`Error HTTP: ${response.status}`);
        }
        
        // GITHUB DEVUELVE EL ARRAY DIRECTO (No usamos .value)
        baseDatosCompleta = await response.json(); 

        console.log("Datos cargados:", baseDatosCompleta);

        // Llenar lista de clientes
        const clientesUnicos = [...new Set(baseDatosCompleta.map(item => item.cliente))];
        const ddlClientes = document.getElementById("ddlClientes");
        
        // Limpiamos y llenamos
        ddlClientes.innerHTML = '<option value="">-- Seleccione Cliente --</option>';
        
        clientesUnicos.forEach(cliente => {
            let opt = document.createElement("option");
            opt.value = cliente;
            opt.textContent = cliente;
            ddlClientes.appendChild(opt);
        });

    } catch (error) {
        console.error("Error crítico cargando datos:", error);
        // Mostramos el error en el dropdown para saber qué pasa
        const ddl = document.getElementById("ddlClientes");
        if(ddl) ddl.innerHTML = '<option>Error de Conexión</option>';
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