/* global Office */

const db = [
    { id: "7560", nombre: "Estudio Pila ROM", cliente: "CODELCO" },
    { id: "8890", nombre: "Ingeniería Tranque", cliente: "ANGLO" }
];

let proyectoActual = null;

Office.onReady(() => {
    document.getElementById("btnSearch").onclick = buscar;
});

function buscar() {
    const val = document.getElementById("inputSearch").value;
    const found = db.find(p => p.id === val);
    
    // Elementos a mostrar/ocultar
    const infoBox = document.getElementById("infoProyecto");
    const plantillasBox = document.getElementById("seccionPlantillas");

    if(found) {
        proyectoActual = found;
        document.getElementById("lblNombre").textContent = found.nombre;
        document.getElementById("lblCliente").textContent = found.cliente;
        
        // Quitamos la clase 'oculto' para mostrar
        infoBox.classList.remove("oculto");
        plantillasBox.classList.remove("oculto");
    } else {
        // Agregamos la clase 'oculto' para esconder
        infoBox.classList.add("oculto");
        plantillasBox.classList.add("oculto");
        alert("Proyecto no encontrado");
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