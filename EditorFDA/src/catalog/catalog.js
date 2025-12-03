/* global Office, Word */

let baseDatosCompleta = [];
let proyectoActual = null;

// RUTAS
const urlFuenteDatos = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/proyectos.json";
// Asumo esta ruta basado en tu JSON. Ajustala si tus plantillas están en otra carpeta.
const urlBasePlantillas = "https://basmon123.github.io/Web-Word/EditorFDA/src/templates/"; 

Office.onReady(async () => {
    await cargarDatosIniciales();
    
    const ddlClientes = document.getElementById("ddlClientes");
    const ddlProyectos = document.getElementById("ddlProyectos");

    if(ddlClientes) ddlClientes.onchange = filtrarProyectos;
    if(ddlProyectos) ddlProyectos.onchange = seleccionarProyecto;
});

// ... (MANTÉN TUS FUNCIONES cargarDatosIniciales, filtrarProyectos, seleccionarProyecto, ocultarDetalles, setText IGUAL QUE ANTES) ...

// --- NUEVA LÓGICA PARA ABRIR LA PLANTILLA ---

window.seleccionarPlantilla = async function(nombreArchivo) {
    // nombreArchivo debe ser el nombre real, ej: "InformeAvance.docx"
    // Si tus botones envían "informe", necesitas mapearlo al nombre del archivo aquí.
    
    if(!proyectoActual) {
        console.log("No hay proyecto seleccionado");
        return;
    }

    try {
        // 1. Definir qué archivo vamos a buscar
        // Si tu botón HTML dice onclick="seleccionarPlantilla('Informe.docx')", usa 'nombreArchivo' directo.
        // Si dice onclick="seleccionarPlantilla('informe')", usa un switch para elegir el archivo:
        let archivoReal = nombreArchivo;
        if (!nombreArchivo.includes(".docx")) {
             archivoReal = nombreArchivo + ".docx"; // Pequeña ayuda por si olvidas la extensión
        }

        const urlCompleta = urlBasePlantillas + archivoReal;
        console.log("Descargando plantilla de:", urlCompleta);

        // 2. Descargar el archivo desde GitHub/Servidor
        const response = await fetch(urlCompleta);
        if (!response.ok) throw new Error("No se pudo descargar la plantilla desde: " + urlCompleta);
        
        const blob = await response.blob();

        // 3. Convertir a Base64 para que Word lo entienda
        const base64 = await getBase64(blob);
        
        // El base64 viene con el encabezado "data:application/...", hay que quitarlo para Word
        const base64Limpio = base64.split(',')[1];

        // 4. Ordenar a Word que cree un nuevo documento con ese Base64
        await Word.run(async (context) => {
            const newDoc = context.application.createDocument(base64Limpio);
            newDoc.open(); // Esto abre la plantilla en una ventana nueva de Word
            await context.sync();
        });

    } catch (error) {
        console.error("Error abriendo plantilla:", error);
        // Opcional: Mostrar error en pantalla
        setText("lblNombre", "Error: " + error.message); 
    }
}

// Función auxiliar necesaria para convertir el archivo descargado
function getBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}