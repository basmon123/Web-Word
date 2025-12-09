/* global document, Office, Word, fetch */

let datosProyectoActual = {};

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // 1. Configurar botones
        const btnAct = document.getElementById("btnActualizarTitulo");
        if(btnAct) btnAct.onclick = actualizarTituloDocumento;
        
        const btnVolver = document.getElementById("btnVolver");
        if(btnVolver) btnVolver.onclick = mostrarCatalogo;

        // 2. Cargar el catálogo
        iniciarCatalogo();
    }
});

// URL DE TU JSON (Asegúrate que tu repo sea PÚBLICO y el nombre del archivo sea exacto)
const URL_JSON = "https://raw.githubusercontent.com/basmon123/templates/main/data.json"; 

// --- A. LÓGICA DEL CATÁLOGO ---

async function iniciarCatalogo() {
    try {
        console.log("Intentando descargar catálogo desde GitHub...");
        const response = await fetch(URL_JSON);
        
        // Verificamos si la respuesta fue exitosa (código 200)
        if (!response.ok) {
            throw new Error(`Error de red: ${response.status}`);
        }

        const data = await response.json();
        const proyectos = data.projects || data; 
        renderizarTarjetas(proyectos);

    } catch (error) {
        console.warn("⚠️ No se pudo cargar desde GitHub. Usando datos locales de prueba.", error);
        // SI FALLA GITHUB, CARGAMOS ESTOS DATOS DE RESPALDO:
        usarDatosLocales();
    }
}

function usarDatosLocales() {
    const proyectosPrueba = [
        {
            "NombreProyecto": "Proyecto Demo Local 1",
            "Cliente": "Cliente Prueba",
            "Division": "División Norte",
            "TipoServicio": "Ingeniería Básica",
            "NumeroContrato": "4600-TEST-01",
            "NumeroAPI": "API-001",
            "NombreDoc": "Informe de Inicio"
        },
        {
            "NombreProyecto": "Estudio de Suelos",
            "Cliente": "Minera Ejemplo",
            "Division": "División Sur",
            "TipoServicio": "Geotecnia",
            "NumeroContrato": "5500-GEO-99",
            "NumeroAPI": "API-GEO",
            "NombreDoc": "Reporte Técnico"
        }
    ];
    renderizarTarjetas(proyectosPrueba);
    
    // Mostramos un aviso visual de que estamos offline/local
    const contenedor = document.getElementById("contenedor-tarjetas");
    const aviso = document.createElement("p");
    aviso.style.color = "orange";
    aviso.style.fontSize = "11px";
    aviso.textContent = "Nota: Mostrando datos de prueba (Error conectando a GitHub).";
    contenedor.insertBefore(aviso, contenedor.firstChild);
}

function renderizarTarjetas(listaProyectos) {
    const contenedor = document.getElementById("contenedor-tarjetas");
    if(!contenedor) return;
    
    contenedor.innerHTML = ""; // Limpiar mensaje de "Cargando..."

    listaProyectos.forEach(proyecto => {
        const card = document.createElement("div");
        card.className = "card";
        card.innerHTML = `
            <h4>${proyecto.NombreProyecto || "Sin Nombre"}</h4>
            <p><strong>Cliente:</strong> ${proyecto.Cliente || "-"}</p>
            <p><strong>División:</strong> ${proyecto.Division || "-"}</p>
        `;

        card.onclick = () => {
            seleccionarProyecto(proyecto);
        };

        contenedor.appendChild(card);
    });
}

async function seleccionarProyecto(proyecto) {
    // 1. Aquí iría la lógica para abrir la plantilla si tuvieras la URL
    console.log("Proyecto seleccionado:", proyecto.NombreProyecto);

    // 2. Cargar datos en la Vista Detalle
    await cargarDatosEnTaskpane(proyecto);

    // 3. Cambiar de pantalla
    mostrarDetalle();
}

// --- B. NAVEGACIÓN ENTRE VISTAS ---

function mostrarDetalle() {
    document.getElementById("vista-catalogo").classList.add("oculto");
    document.getElementById("vista-detalle").classList.remove("oculto");
}

function mostrarCatalogo() {
    document.getElementById("vista-catalogo").classList.remove("oculto");
    document.getElementById("vista-detalle").classList.add("oculto");
    
    const msg = document.getElementById("mensajeEstado");
    if(msg) msg.textContent = "";
}

// --- C. LÓGICA DEL DETALLE ---

window.cargarDatosEnTaskpane = async function(datos) {
    datosProyectoActual = datos;

    setText("lblCliente", datos.Cliente);
    setText("lblDivision", datos.Division);
    setText("lblServicio", datos.TipoServicio);
    setText("lblContrato", datos.NumeroContrato);
    setText("lblApi", datos.NumeroAPI);
    setText("lblProyecto", datos.NombreProyecto);

    const inputTitulo = document.getElementById("txtTituloDoc");
    if(inputTitulo) inputTitulo.value = datos.NombreDoc || "";

    await escribirDatosBaseEnWord(datos);
};

async function actualizarTituloDocumento() {
    const msgLabel = document.getElementById("mensajeEstado");
    const nuevoTitulo = document.getElementById("txtTituloDoc").value;

    if (!nuevoTitulo) {
        if (msgLabel) msgLabel.textContent = "⚠️ El título está vacío.";
        return;
    }
    
    if (msgLabel) msgLabel.textContent = "Actualizando...";

    await Word.run(async (context) => {
        const controls = context.document.contentControls.getByTag("ccNombreDoc");
        controls.load("items");
        await context.sync();
        
        let count = 0;
        if (controls.items.length > 0) {
             controls.items.forEach(cc => {
                 cc.insertText(nuevoTitulo, "Replace");
                 count++;
             });
        }
        await context.sync();
        
        if (msgLabel) msgLabel.textContent = count > 0 ? "✅ Título actualizado." : "⚠️ No se encontró 'ccNombreDoc' en el Word.";
    });
}

async function escribirDatosBaseEnWord(datos) {
    await Word.run(async (context) => {
        const tagsMapa = [
            { t: "ccCliente",              v: datos.Cliente }, 
            { t: "ccDivisión",             v: datos.Division },
            { t: "ccServicios",            v: datos.TipoServicio },
            { t: "ccContrato",             v: datos.NumeroContrato },
            { t: "ccAPI",                  v: datos.NumeroAPI },
            { t: "ccProyecto",             v: datos.NombreProyecto }
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
    if (el) el.textContent = val || "-";
}