/* global Office */

let baseDatosCompleta = [];
let proyectoActual = null;

Office.onReady(async () => {
    await cargarDatosIniciales();
    
    // Eventos de cambio
    document.getElementById("ddlClientes").onchange = filtrarProyectos;
    document.getElementById("ddlProyectos").onchange = seleccionarProyecto;
});

async function cargarDatosIniciales() {
    try {
        // Cargar JSON
        const response = await fetch("https://basmon123.github.io/Web-Word/EditorFDA/src/data/proyectos.json");
        baseDatosCompleta = await response.json();

        // Obtener Clientes Únicos (Para no repetir "Codelco" 20 veces en la lista)
        const clientesUnicos = [...new Set(baseDatosCompleta.map(item => item.cliente))];
        
        // Llenar el Dropdown de Clientes
        const ddlClientes = document.getElementById("ddlClientes");
        ddlClientes.innerHTML = '<option value="">-- Seleccione Cliente --</option>';
        
        clientesUnicos.forEach(cliente => {
            let opt = document.createElement("option");
            opt.value = cliente;
            opt.textContent = cliente;
            ddlClientes.appendChild(opt);
        });

    } catch (error) {
        console.error("Error cargando datos:", error);
    }
}

function filtrarProyectos() {
    const clienteSeleccionado = document.getElementById("ddlClientes").value;
    const ddlProyectos = document.getElementById("ddlProyectos");
    
    // Resetear lista de proyectos
    ddlProyectos.innerHTML = '<option value="">-- Seleccione Proyecto --</option>';
    document.getElementById("seccionPlantillas").classList.add("oculto");
    document.getElementById("infoProyecto").classList.add("oculto");

    if (!clienteSeleccionado) {
        ddlProyectos.disabled = true;
        return;
    }

    // Filtrar: Dame solo los proyectos de este cliente
    const proyectosFiltrados = baseDatosCompleta.filter(p => p.cliente === clienteSeleccionado);

    // Llenar Dropdown
    proyectosFiltrados.forEach(p => {
        let opt = document.createElement("option");
        opt.value = p.id; // El valor es el ID (7560)
        opt.textContent = p.id + " - " + p.nombre; // Lo que se ve
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

    // Buscar el objeto completo del proyecto
    proyectoActual = baseDatosCompleta.find(p => p.id === idProyecto);

    // Mostrar info
    document.getElementById("lblNombre").textContent = proyectoActual.nombre;
    document.getElementById("lblAPI").textContent = "API: " + proyectoActual.api;
    document.getElementById("infoProyecto").classList.remove("oculto");
    
    // Mostrar Plantillas
    document.getElementById("seccionPlantillas").classList.remove("oculto");
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