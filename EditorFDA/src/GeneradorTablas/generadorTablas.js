/* global Office, document */

Office.onReady(() => {
    // Eventos para actualizar la vista previa
    document.getElementById("txtFilas").oninput = actualizarPreview;
    document.getElementById("txtCols").oninput = actualizarPreview;
    
    // Evento del botón Cargar
    document.getElementById("btnInsertar").onclick = enviarDatosAWord;
    
    actualizarPreview(); // Primera carga
});

function actualizarPreview() {
    const f = document.getElementById("txtFilas").value;
    const c = document.getElementById("txtCols").value;
    const tabla = document.getElementById("tablaPreview");
    
    tabla.innerHTML = "";
    for(let i=0; i<f; i++){
        let row = tabla.insertRow();
        for(let j=0; j<c; j++){
            let cell = row.insertCell();
            cell.innerHTML = "&nbsp;"; // Espacio vacío
        }
    }
}

function enviarDatosAWord() {
    // 1. Recopilar datos
    const config = {
        filas: document.getElementById("txtFilas").value,
        columnas: document.getElementById("txtCols").value,
        estilo: document.getElementById("selEstilo").value
    };

    // 2. Enviar mensaje a la ventana padre (commands.js)
    Office.context.ui.messageParent(JSON.stringify(config));
}