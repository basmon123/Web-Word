/* global Office */

let baseDatosCompleta = [];
let proyectoActual = null;

// URL FIJA DE GITHUB
const urlFuenteDatos = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/proyectos.json";

Office.onReady(async () => {
    // 1. Cargamos datos
    await cargarDatosIniciales();
    
    // 2. Eventos de los selectores
    // Verificamos que existan los elementos antes de asignar eventos para evitar errores si el HTML cambia
    const ddlClientes = document.getElementById("ddlClientes");
    const ddlProyectos = document.getElementById("ddlProyectos");
    const btnSearch = document.getElementById("btnSearch");

    if(ddlClientes) ddlClientes.onchange = filtrarProyectos;
    if(ddlProyectos) ddlProyectos.onchange = seleccionarProyecto;
    if(btnSearch) btnSearch.onclick = buscar;
});

async function cargarDatosIniciales() {
    try {
        console.log("Intentando cargar datos desde:", urlFuenteDatos);
        
        // "Engañar" a la caché
        const response = await fetch(urlFuenteDatos + "?t=" + new Date().getTime());

        if (!response.ok) {
            throw new Error("Error HTTP " + response.status);
        }

        const data = await response.json();
        console.log("Datos recibidos:", data);

        // --- LÓGICA DE DETECCIÓN DE ESTRUCTURA ---
        let listaParaUsar = [];

        if (data.body && Array.isArray(data.body)) {
            listaParaUsar = data.body; // Caso Power Automate
        } else if (Array.isArray(data)) {
            listaParaUsar = data; // Caso Array directo
        } else {
            console.error("Formato JSON no reconocido:", data);
            return;
        }

        // --- MAPEO DE DATOS ---
        const proyectosFormateados = listaParaUsar.map(item => {
            return {
                id: item.id || item.Title || item.ID, 
                nombre: item.nombre || item.NombreProyecto, 
                cliente: item.cliente || item.Cliente,
                division: item.division || item.Division,
                contrato: item.contrato || item.Contrato
            };
        });

        console.log("Proyectos procesados:", proyectosFormateados);

        // --- CORRECCIÓN CRÍTICA 1: Actualizar la variable global ---
        baseDatosCompleta = proyectosFormateados;

        // --- CORRECCIÓN CRÍTICA 2: Llenar el primer Dropdown (Clientes) ---
        llenarDropdownClientes();

    } catch (error) {
        console.error("Error crítico cargando datos:", error);
    }
}

// Nueva función para poblar el select de Clientes (ddlClientes)
function llenarDropdownClientes() {
    const ddlClientes = document.getElementById("ddlClientes");
    if (!ddlClientes) return;

    ddlClientes.innerHTML = '<option value="">-- Seleccione Cliente --</option>';

    // Obtener clientes únicos usando un Set
    const clientesUnicos = [...new Set(baseDatosCompleta.map(p => p.cliente))].sort();

    clientesUnicos.forEach(cliente => {
        if (cliente) { // Evitar nulos
            let opt = document.createElement("option");
            opt.value = cliente;
            opt.textContent = cliente;
            ddlClientes.appendChild(opt);
        }
    });
}

function filtrarProyectos() {
    const clienteSeleccionado = document.getElementById("ddlClientes").value;
    const ddlProyectos = document.getElementById("ddlProyectos");
    
    // Resetear el segundo dropdown
    ddlProyectos.innerHTML = '<option value="">-- Seleccione Proyecto --</option>';
    
    // Ocultar paneles de info
    const secPlantillas = document.getElementById("seccionPlantillas");
    const infoProy = document.getElementById("infoProyecto");
    if(secPlantillas) secPlantillas.classList.add("oculto");
    if(infoProy) infoProy.classList.add("oculto");

    if (!clienteSeleccionado) {
        ddlProyectos.disabled = true;
        return;
    }

    // Filtramos usando la variable global que AHORA SÍ tiene datos
    const proyectosFiltrados = baseDatosCompleta.filter(p => p.cliente === clienteSeleccionado);

    proyectosFiltrados.forEach(p => {
        let opt = document.createElement("option");
        opt.value = p.id;
        opt.textContent = `${p.nombre} (${p.contrato || 'S/C'})`; // Agregué contrato para mejor visualización
        ddlProyectos.appendChild(opt);
    });

    ddlProyectos.disabled = false;
}

function seleccionarProyecto() {
    const idProyecto = document.getElementById("ddlProyectos").value;
    
    const secPlantillas = document.getElementById("seccionPlantillas");
    const infoProy = document.getElementById("infoProyecto");

    if (!idProyecto) {
        if(secPlantillas) secPlantillas.classList.add("oculto");
        return;
    }

    proyectoActual = baseDatosCompleta.find(p => p.id === idProyecto);

    if (proyectoActual) {
        // Llenar etiquetas visuales (asegúrate que estos IDs existan en tu HTML)
        const lblNombre = document.getElementById("lblNombre");
        if (lblNombre) lblNombre.textContent = proyectoActual.nombre;
        
        // Mostramos las secciones
        if(infoProy) infoProy.classList.remove("oculto");
        if(secPlantillas) secPlantillas.classList.remove("oculto");
    }
}

// BÚSQUEDA MANUAL
function buscar() {
    const inputSearch = document.getElementById("inputSearch");
    if (!inputSearch) return;

    const val = inputSearch.value.trim(); // Trim para quitar espacios accidentales
    const found = baseDatosCompleta.find(p => p.id === val);
    
    if(found) {
        // Simulamos la selección en cascada
        const ddlClientes = document.getElementById("ddlClientes");
        const ddlProyectos = document.getElementById("ddlProyectos");

        if (ddlClientes) {
            ddlClientes.value = found.cliente;
            filtrarProyectos(); // Esto llena el ddlProyectos
        }
        
        if (ddlProyectos) {
            ddlProyectos.value = found.id;
            seleccionarProyecto(); // Esto muestra la info
        }
    } else {
        // Usar Office.context.ui.displayDialogAsync o un simple alert si no hay otra opción
        console.log("Proyecto no encontrado"); 
        // alert("Proyecto no encontrado"); // Descomentar si quieres alerta visual
    }
}

// Función global para ser llamada desde el HTML si es necesario
window.seleccionarPlantilla = function(tipo) {
    if(!proyectoActual) return;
    const mensaje = {
        accion: "CREAR_DOCUMENTO",
        plantilla: tipo,
        datos: proyectoActual
    };
    // Enviamos mensaje al padre (Taskpane) o procesamos directamente
    console.log("Enviando a Word:", mensaje);
    // Office.context.ui.messageParent(JSON.stringify(mensaje)); // Solo si usas dialogos
    
    // Si estás en el Taskpane directo, aquí llamarías a tu función de Word.run
    // insertarDatosEnDocumento(proyectoActual); 
}