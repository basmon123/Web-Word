/* global document, Office, Word, fetch */

let datosProyectoActual = {};

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // 1. Configurar botones
        document.getElementById("btnActualizarTitulo").onclick = actualizarTituloDocumento;
        document.getElementById("btnVolver").onclick = mostrarCatalogo;

        // 2. Cargar el catálogo al iniciar
        iniciarCatalogo();
    }
});

//URL DE TU JSON EN GITHUB
const URL_JSON = "https://raw.githubusercontent.com/basmon123/templates/main/data.json"; 

// --- A. LÓGICA DEL CATÁLOGO ---

async function iniciarCatalogo() {
    try {
        const response = await fetch(URL_JSON);
        const data = await response.json();
        
        // Asumiendo que tu JSON tiene una propiedad "projects" o es un array directo
        // Ajusta esto según la estructura de tu JSON (ej: data.projects o data)
        const proyectos = data.projects || data; 
        
        renderizarTarjetas(proyectos);
    } catch (error) {
        console.error("Error cargando catálogo:", error);
        document.getElementById("contenedor-tarjetas").innerHTML = "<p style='color:red'>Error al cargar proyectos.</p>";
    }
}

function renderizarTarjetas(listaProyectos) {
    const contenedor = document.getElementById("contenedor-tarjetas");
    contenedor.innerHTML = ""; // Limpiar

    listaProyectos.forEach(proyecto => {
        // Crear tarjeta HTML
        const card = document.createElement("div");
        card.className = "card";
        card.innerHTML = `
            <h4>${proyecto.NombreProyecto || "Sin Nombre"}</h4>
            <p><strong>Cliente:</strong> ${proyecto.Cliente || "-"}</p>
            <p><strong>División:</strong> ${proyecto.Division || "-"}</p>
        `;

        // EVENTO CLAVE: AL HACER CLIC EN LA TARJETA
        card.onclick = () => {
            seleccionarProyecto(proyecto);
        };

        contenedor.appendChild(card);
    });
}

async function seleccionarProyecto(proyecto) {
    // 1. Abrir el documento (Lógica que ya tenías, aquí simplificada)
    // Si tu JSON tiene una URL de plantilla, úsala aquí.
    if (proyecto.urlPlantilla) {
        // createFromTemplate(proyecto.urlPlantilla)... (Tu lógica de apertura)
        console.log("Abriendo plantilla: " + proyecto.urlPlantilla);
        
        // NOTA: Si abres un documento nuevo, el Add-in puede recargarse.
        // Si el Add-in se recarga, perderemos el estado. 
        // Para Add-ins persistentes, asumimos que trabajas sobre el doc actual 
        // o que el usuario ya abrió el doc y solo selecciona los datos.
    }

    // 2. Cargar datos en la Vista Detalle
    await cargarDatosEnTaskpane(proyecto);

    // 3. Cambiar de pantalla
    mostrarDetalle();
}

// --- B. NAVEGACIÓN ENTRE VISTAS ---

// --- B. NAVEGACIÓN ENTRE VISTAS ---

function mostrarDetalle() {
    // Ocultar catálogo agregando la clase
    document.getElementById("vista-catalogo").classList.add("oculto");
    // Mostrar detalle quitando la clase
    document.getElementById("vista-detalle").classList.remove("oculto");
}

function mostrarCatalogo() {
    // Mostrar catálogo quitando la clase
    document.getElementById("vista-catalogo").classList.remove("oculto");
    // Ocultar detalle agregando la clase
    document.getElementById("vista-detalle").classList.add("oculto");
    
    // Opcional: Limpiar mensaje de estado al volver
    const msg = document.getElementById("mensajeEstado");
    if(msg) msg.textContent = "";
}
// --- C. LÓGICA DEL DETALLE (Tu código nuevo) ---

window.cargarDatosEnTaskpane = async function(datos) {
    console.log("Cargando datos...", datos);
    datosProyectoActual = datos;

    // Llenar textos (SPAN)
    setText("lblCliente", datos.Cliente);
    setText("lblDivision", datos.Division);
    setText("lblServicio", datos.TipoServicio);
    setText("lblContrato", datos.NumeroContrato);
    setText("lblApi", datos.NumeroAPI);
    setText("lblProyecto", datos.NombreProyecto);

    // Llenar input editable
    const inputTitulo = document.getElementById("txtTituloDoc");
    if(inputTitulo) inputTitulo.value = datos.NombreDoc || "";

    // Escribir automáticamente en Word los datos base
    await escribirDatosBaseEnWord(datos);
};

// ... (Aquí van tus funciones actualizarTituloDocumento y escribirDatosBaseEnWord tal cual las tenías antes) ...

async function actualizarTituloDocumento() {
    // ... Tu lógica de actualizar título ...
    const nuevoTitulo = document.getElementById("txtTituloDoc").value;
    await Word.run(async (context) => {
        const controls = context.document.contentControls.getByTag("ccNombreDoc");
        controls.load("items");
        await context.sync();
        if (controls.items.length > 0) {
             controls.items[0].insertText(nuevoTitulo, "Replace");
        }
    });
    document.getElementById("mensajeEstado").textContent = "Título actualizado.";
}

async function escribirDatosBaseEnWord(datos) {
     // ... Tu lógica de escribir los datos duros ...
     // Copia aquí la función escribirDatosBaseEnWord que te di en la respuesta anterior
}

function setText(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val || "-";
}