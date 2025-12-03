/* global Office */



let baseDatosCompleta = [];

let proyectoActual = null;



// URL FIJA DE GITHUB

const urlFuenteDatos = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/proyectos.json";

// ANTES: .../templates/" + datosProyecto.id + "/"...
// AHORA: Usamos .carpeta_plantilla (Ej: CODELCO)
const urlPlantilla = "https://basmon123.github.io/Web-Word/EditorFDA/src/templates/" + datosProyecto.carpeta_plantilla + "/" + nombreArchivo;

Office.onReady(async () => {

    await cargarDatosIniciales();

   

    // Eventos de los selectores

    const ddlClientes = document.getElementById("ddlClientes");

    const ddlProyectos = document.getElementById("ddlProyectos");



    if(ddlClientes) ddlClientes.onchange = filtrarProyectos;

    if(ddlProyectos) ddlProyectos.onchange = seleccionarProyecto;

});



async function cargarDatosIniciales() {

    try {

        // Cache-busting

        const response = await fetch(urlFuenteDatos + "?t=" + new Date().getTime());

        const data = await response.json();



        // Detección de estructura (Power Automate vs Array directo)

        let lista = [];

        if (data.body && Array.isArray(data.body)) {

            lista = data.body;

        } else if (Array.isArray(data)) {

            lista = data;

        }



        // Mapeo de datos

        baseDatosCompleta = lista.map(item => ({

            id: item.id || item.Title || item.ID,

            nombre: item.nombre || item.NombreProyecto,

            cliente: item.cliente || item.Cliente,

            division: item.division || item.Division,

            contrato: item.contrato || item.Contrato,

            api: item.api || item.API,

            carpeta_plantilla: item.carpeta_plantilla || "General"

        }));



        // Llenar Dropdown Clientes

        const ddlClientes = document.getElementById("ddlClientes");

        ddlClientes.innerHTML = '<option value="">-- Seleccione Cliente --</option>';

       

        // Clientes únicos y ordenados

        const clientesUnicos = [...new Set(baseDatosCompleta.map(p => p.cliente))].sort();

       

        clientesUnicos.forEach(c => {

            if(c) {

                let opt = document.createElement("option");

                opt.value = c;

                opt.textContent = c;

                ddlClientes.appendChild(opt);

            }

        });



    } catch (error) {

        console.error("Error cargando datos:", error);

        document.getElementById("ddlClientes").innerHTML = '<option>Error de conexión</option>';

    }

}



function filtrarProyectos() {

    const clienteSel = document.getElementById("ddlClientes").value;

    const ddlProyectos = document.getElementById("ddlProyectos");

   

    // Resetear segunda lista y ocultar todo

    ddlProyectos.innerHTML = '<option value="">-- Seleccione N° --</option>';

    ocultarDetalles();



    if (!clienteSel) {

        ddlProyectos.disabled = true;

        return;

    }



    // Filtrar proyectos del cliente

    const filtrados = baseDatosCompleta.filter(p => p.cliente === clienteSel);



    filtrados.forEach(p => {

        let opt = document.createElement("option");

        // AQUÍ ESTÁ LO QUE PEDISTE:

        // Text: ID del Proyecto (Ej: 7560)

        // Value: ID del Proyecto

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



    // Buscar proyecto seleccionado

    proyectoActual = baseDatosCompleta.find(p => p.id === idProyecto);



    if (proyectoActual) {

        // Llenar la ficha de detalles

        setText("lblNombre", proyectoActual.nombre);

        setText("lblCliente", proyectoActual.cliente);

        setText("lblDivision", proyectoActual.division);

        setText("lblContrato", proyectoActual.contrato);

        setText("lblAPI", proyectoActual.api);



        // Mostrar secciones

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



// Función que envía la orden a Word

window.seleccionarPlantilla = function(tipo) {

    if(!proyectoActual) return;
    const mensaje = {

        accion: "CREAR_DOCUMENTO",

        plantilla: tipo,

        datos: proyectoActual

    };

    Office.context.ui.messageParent(JSON.stringify(mensaje));

}