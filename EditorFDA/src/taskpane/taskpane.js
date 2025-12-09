/* global document, Office, Word */

// Variable global para guardar los datos cargados temporalmente (opcional)
let datosProyectoActual = {};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Asignar el evento al NUEVO botón de actualizar título
    const btn = document.getElementById("btnActualizarTitulo");
    if (btn) btn.onclick = actualizarTituloDocumento;
  }
});

/**
 * ---------------------------------------------------------
 * 1. FUNCIÓN PARA CARGAR DATOS (Lectura desde SharePoint)
 * ---------------------------------------------------------
 * Esta función la debes llamar desde tu lógica de Catálogo/SharePoint
 * pasándole el objeto JSON del proyecto seleccionado.
 */
window.cargarDatosEnTaskpane = async function(datos) {
    try {
        // A. Guardamos datos en variable global por si se necesitan luego
        datosProyectoActual = datos;

        // B. Llenamos la UI (Los SPAN de solo lectura)
        setText("lblCliente", datos.Cliente);
        setText("lblDivision", datos.Division);
        setText("lblServicio", datos.TipoServicio);
        setText("lblContrato", datos.NumeroContrato);
        setText("lblApi", datos.NumeroAPI);
        setText("lblProyecto", datos.NombreProyecto);

        // C. Pre-llenar el input del Título si viene dato, si no, dejar vacío
        const inputTitulo = document.getElementById("txtTituloDoc");
        if(inputTitulo && datos.NombreDoc) {
            inputTitulo.value = datos.NombreDoc;
        }

        // D. Escribir AUTOMÁTICAMENTE en el Word los datos "Duros" del proyecto
        //    (Cliente, Contrato, etc.) para que el usuario no tenga que hacerlo.
        await escribirDatosBaseEnWord(datos);

    } catch (error) {
        console.error("Error al cargar datos:", error);
    }
};

/**
 * ---------------------------------------------------------
 * 2. FUNCIÓN DE ESCRITURA MANUAL (Solo Título)
 * ---------------------------------------------------------
 * Se ejecuta al dar clic en "Actualizar Título"
 */
async function actualizarTituloDocumento() {
  try {
    const msgLabel = document.getElementById("mensajeEstado");
    const nuevoTitulo = document.getElementById("txtTituloDoc").value;

    if (!nuevoTitulo) {
        if (msgLabel) msgLabel.textContent = "⚠️ El título está vacío.";
        return;
    }

    if (msgLabel) msgLabel.textContent = "Actualizando título...";

    await Word.run(async (context) => {
      // Solo buscamos el Content Control del título
      // Asegúrate que en Word el Tag sea "ccNombreDoc"
      const controls = context.document.contentControls.getByTag("ccNombreDoc");
      controls.load("items");
      
      await context.sync();

      let count = 0;
      if (controls.items.length > 0) {
        // Insertar texto en todos los controles que tengan ese Tag
        controls.items.forEach((cc) => {
            cc.insertText(nuevoTitulo, "Replace");
            count++;
        });
      }

      await context.sync();
      
      if (msgLabel) {
          msgLabel.textContent = count > 0 
            ? "✅ Título actualizado." 
            : "⚠️ No se encontró el control 'ccNombreDoc' en el Word.";
      }
    });

  } catch (error) {
    console.error(error);
    const msgLabel = document.getElementById("mensajeEstado");
    if (msgLabel) msgLabel.textContent = "❌ Error: " + error.message;
  }
}

/**
 * ---------------------------------------------------------
 * 3. AUXILIAR: Escribe los datos fijos de SharePoint en Word
 * ---------------------------------------------------------
 */
async function escribirDatosBaseEnWord(datos) {
    await Word.run(async (context) => {
        // Mapeo de datos JSON -> Tags de Content Control en Word
        // Ajusta los nombres de la derecha (datos.X) según tu JSON de SharePoint
        const tagsMapa = [
            { t: "ccCliente",              v: datos.Cliente }, 
            { t: "ccCliente_encabezado",   v: datos.Cliente },
            { t: "ccDivisión",             v: datos.Division },
            { t: "ccD_encabezado",         v: datos.Division },
            { t: "ccServicios",            v: datos.TipoServicio },
            { t: "ccContrato",             v: datos.NumeroContrato },
            { t: "ccAPI",                  v: datos.NumeroAPI },
            { t: "ccProyecto",             v: datos.NombreProyecto },
            { t: "ccNProyecto_Encabezado", v: datos.NombreProyecto }
            // Nota: El título (ccNombreDoc) no lo actualizamos aquí, 
            // lo dejamos para el botón manual o si quieres forzarlo, agrégalo.
        ];

        for (let item of tagsMapa) {
            if (!item.v) continue; // Si el dato es null o vacío, saltamos

            let ccs = context.document.contentControls.getByTag(item.t);
            ccs.load("items");
            await context.sync();

            if (ccs.items.length > 0) {
                ccs.items.forEach(cc => {
                    cc.insertText(item.v, "Replace");
                });
            }
        }
        await context.sync();
    });
}

// Helper simple para poner texto en los labels
function setText(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val || "-";
}