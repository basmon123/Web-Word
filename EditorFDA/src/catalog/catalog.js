/* global Office */

let proyectoActual = null;

Office.onReady(() => {
    document.getElementById("btnSearch").onclick = buscar;
});

async function buscar() {
    const val = document.getElementById("inputSearch").value;
    const infoBox = document.getElementById("infoProyecto");
    const plantillasBox = document.getElementById("seccionPlantillas");

    // 1. CARGAR DATOS DESDE LA NUBE (GITHUB)
    // Usamos la ruta relativa al servidor
    try {
        const response = await fetch("../../data/proyectos.json"); // Sube 2 niveles para encontrar data
        const db = await response.json(); // Convierte el texto a objetos

        // 2. BUSCAR EN LA LISTA DESCARGADA
        const found = db.find(p => p.id === val);
        
        if(found) {
            proyectoActual = found;
            document.getElementById("lblNombre").textContent = found.nombre;
            document.getElementById("lblCliente").textContent = found.cliente;
            
            infoBox.classList.remove("oculto");
            plantillasBox.classList.remove("oculto");
        } else {
            infoBox.classList.add("oculto");
            plantillasBox.classList.add("oculto");
            // Usamos un mensaje en pantalla en vez de alert para ser más elegantes
            document.getElementById("lblNombre").textContent = "No encontrado";
            infoBox.classList.remove("oculto");
        }

    } catch (error) {
        console.error("Error cargando base de datos:", error);
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