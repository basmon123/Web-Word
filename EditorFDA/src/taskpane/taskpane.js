/* global document, Office, Word */

// 1. Configuración Inicial
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Solo intentamos asignar el clic si el botón realmente existe en el HTML
    const btn = document.getElementById("btnGenerar");
    if (btn) {
      btn.onclick = run;
    } else {
      console.error("No encontré el botón 'btnGenerar' en el HTML.");
    }
  }
});

// 2. La Lógica Principal
async function run() {
  try {
    // A. Capturar datos (Usamos 'value' con seguridad)
    const getVal = (id) => document.getElementById(id) ? document.getElementById(id).value : "";
    
    const vCliente   = getVal("inCliente");
    const vDivision  = getVal("inDivision");
    const vProyecto  = getVal("inProyecto");
    const vContrato  = getVal("inContrato");
    const vAPI       = getVal("inAPI");
    const vServicios = getVal("inServicios");
    const vNombreDoc = getVal("inNombreDoc");
    const vCodigo    = getVal("inCodigo");
    const vRevision  = getVal("inRevision");

    // Mensaje de estado
    const msgLabel = document.getElementById("mensajeEstado");
    if (msgLabel) msgLabel.textContent = "Procesando...";

    await Word.run(async (context) => {
      
      // B. Mapa de Tags vs Valores
      const mapaDeTags = [
        { tag: "ccCliente",       valor: vCliente },
        { tag: "ccDivisión",      valor: vDivision },
        { tag: "ccServicios",     valor: vServicios },
        { tag: "ccContrato",      valor: vContrato },
        { tag: "ccAPI",           valor: vAPI },
        { tag: "ccProyecto",      valor: vProyecto },
        { tag: "ccNombreDoc",    valor: vNombreDoc },
        // Encabezados
        { tag: "ccCliente_encabezado",   valor: vCliente },
        { tag: "ccD_encabezado",         valor: vDivision },
        { tag: "ccNProyecto_Encabezado", valor: vProyecto },
        { tag: "ccCodigo",               valor: vCodigo },
        { tag: "ccRevision",             valor: vRevision }
      ];

      // C. Buscar y Reemplazar
      let contadores = 0;
      
      for (let item of mapaDeTags) {
        // Buscamos los controles por su etiqueta
        let ccs = context.document.contentControls.getByTag(item.tag);
        ccs.load("items");
        await context.sync(); // Sincronización parcial para leer

        if (ccs.items.length > 0) {
           for (let cc of ccs.items) {
             // Escribimos el dato
             cc.insertText(item.valor, "Replace");
             contadores++;
           }
        }
      }

      await context.sync(); // Guardado final
      
      if (msgLabel) msgLabel.textContent = "¡Listo! " + contadores + " campos actualizados.";
      
    });
  } catch (error) {
    console.error(error);
    const msgLabel = document.getElementById("mensajeEstado");
    if (msgLabel) msgLabel.textContent = "Error: " + error.message;
  }
}

// --- NUEVAS FUNCIONES DE LOS BOTONES (Comandos) ---
// Botón 1: Limpieza FDA (MEJORADO)
async function limpiarFormato(event) {
  try {
    await Word.run(async (context) => {
      // 1. Obtenemos lo que el usuario seleccionó
      const selection = context.document.getSelection();
      
      // 2. Cargamos propiedades (opcional, buena práctica)
      context.load(selection, 'font');

      // 3. Aplicamos Estilo FDA
      // Usamos "#000000" (Hex) en vez de "black" para evitar problemas de idioma
      selection.font.name = "Arial";
      selection.font.size = 11;
      selection.font.color = "#000000"; 
      
      // Alineación Justificada
      selection.paragraphFormat.alignment = "Justified"; 

      // 4. Sincronizamos con Word
      await context.sync();
    });
  } catch (error) {
    // Si falla, esto lo verás en la consola del inspector
    console.error("Error al dar formato:", error);
  }
  
  // Avisamos a Word que terminamos
  if (event) event.completed();
}

// Botón 2: Insertar Fecha
async function insertarFecha(event) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    
    // Fecha de hoy en formato local (ej: 02/12/2025)
    const fechaHoy = new Date().toLocaleDateString();
    
    selection.insertText(fechaHoy, "Replace");
    
    await context.sync();
  });
  
  event.completed();
}

// --- REGISTRO DE FUNCIONES (Vital para que el XML las encuentre) ---
// Esto conecta el nombre del XML <FunctionName> con la función de JS
Office.actions.associate("limpiarFormato", limpiarFormato);
Office.actions.associate("insertarFecha", insertarFecha);