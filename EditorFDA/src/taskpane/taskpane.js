/* global document, Office, Word */

// URL DE TU BASE DE DATOS (GitHub)
const urlFuenteDatos = "https://basmon123.github.io/Web-Word/EditorFDA/src/data/proyectos.json";
let baseDatosCompleta = [];

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {
    
    // 1. Configurar botón principal
    const btn = document.getElementById("run"); // Cambié el ID a 'run' en el HTML nuevo
    if (btn) btn.onclick = run;

    // 2. Cargar datos para los selectores
    await cargarDatosIniciales();

    // 3. Configurar eventos de los selectores
    const ddlClients = document.getElementById("client-select");
    const ddlProjects = document.getElementById("project-select");

    if (ddlClients) ddlClients.onchange = filtrarProyectos;
    if (ddlProjects) ddlProjects.onchange = mostrarDetallesProyecto;
  }
});

// --- LÓGICA DE DATOS (NUEVO) ---

async function cargarDatosIniciales() {
    try {
        // Cache-busting para asegurar datos frescos
        const response = await fetch(urlFuenteDatos + "?t=" + new Date().getTime());
        const data = await response.json();

        // Detectar estructura (si viene de Power Automate o directo)
        let lista = [];
        if (data.body && Array.isArray(data.body)) {
            lista = data.body;
        } else if (Array.isArray(data)) {
            lista = data;
        }

        // Mapeo seguro
        baseDatosCompleta = lista.map(item => ({
            id: item.id || item.Title || item.ID,
            nombre: item.nombre || item.NombreProyecto,
            cliente: item.cliente || item.Cliente,
            division: item.division || item.Division,
            contrato: item.contrato || item.Contrato,
            api: item.api || item.API
        }));

        llenarDropdownClientes();

    } catch (error) {
        console.error("Error cargando datos:", error);
    }
}

function llenarDropdownClientes() {
    const ddl = document.getElementById("client-select");
    if(!ddl) return;

    // Obtener clientes únicos y ordenarlos
    const clientes = [...new Set(baseDatosCompleta.map(p => p.cliente))].sort();
    
    clientes.forEach(c => {
        if(c) {
            let opt = document.createElement("option");
            opt.value = c;
            opt.textContent = c;
            ddl.appendChild(opt);
        }
    });
}

function filtrarProyectos() {
    const clienteSel = document.getElementById("client-select").value;
    const ddlProyectos = document.getElementById("project-select");
    
    // Limpiar
    ddlProyectos.innerHTML = '<option value="">-- Selecciona N° --</option>';
    limpiarDetallesVisuales();

    if (!clienteSel) return;

    // Filtrar
    const filtrados = baseDatosCompleta.filter(p => p.cliente === clienteSel);

    filtrados.forEach(p => {
        let opt = document.createElement("option");
        opt.text = p.id; // Mostramos el NÚMERO (ID)
        opt.value = p.id; 
        // Guardamos todo el objeto en el elemento para acceso rápido
        opt.dataset.project = JSON.stringify(p);
        ddlProyectos.appendChild(opt);
    });
}

function mostrarDetallesProyecto() {
    const ddl = document.getElementById("project-select");
    const opt = ddl.options[ddl.selectedIndex];

    if (!opt.value || !opt.dataset.project) {
        limpiarDetallesVisuales();
        return;
    }

    const p = JSON.parse(opt.dataset.project);

    // Llenar los SPAN visuales (Solo lectura)
    setText("lblNombre", p.nombre);
    setText("lblCliente", p.cliente);
    setText("lblDivision", p.division);
    setText("lblContrato", p.contrato);
    setText("lblAPI", p.api);
}

function limpiarDetallesVisuales() {
    setText("lblNombre", "---");
    setText("lblCliente", "---");
    setText("lblDivision", "---");
    setText("lblContrato", "---");
    setText("lblAPI", "---");
}

function setText(id, val) {
    const el = document.getElementById(id);
    if(el) el.textContent = val || "N/A";
}

// --- 1. FUNCIÓN DEL PANEL (GENERADOR / INSERTAR) ---
async function run() {
  try {
    const msgLabel = document.getElementById("mensajeEstado");
    if (msgLabel) msgLabel.textContent = "Procesando...";

    // Obtener datos: Prioridad al proyecto seleccionado, sino buscamos inputs manuales (si existen)
    const ddlProj = document.getElementById("project-select");
    let datosProyecto = {};

    if (ddlProj && ddlProj.value && ddlProj.options[ddlProj.selectedIndex].dataset.project) {
        datosProyecto = JSON.parse(ddlProj.options[ddlProj.selectedIndex].dataset.project);
    }

    // Función auxiliar para leer inputs manuales
    const getVal = (id) => document.getElementById(id) ? document.getElementById(id).value : "";

    // Construimos el objeto final mezclando automático + manual
    const datosFinales = {
        cliente: datosProyecto.cliente || getVal("inCliente"),
        division: datosProyecto.division || getVal("inDivision"),
        proyecto: datosProyecto.nombre || getVal("inProyecto"),
        contrato: datosProyecto.contrato || getVal("inContrato"),
        api: datosProyecto.api || getVal("inAPI"),
        
        // Estos siempre son manuales
        servicios: getVal("inServicios"),
        nombreDoc: getVal("inNombreDoc"),
        codigo: getVal("inCodigo"),
        revision: getVal("inRevision")
    };

    await Word.run(async (context) => {
      let contadores = 0;
      
      const tagsMapa = [
          { t: "ccCliente", v: datosFinales.cliente }, { t: "ccCliente_encabezado", v: datosFinales.cliente },
          { t: "ccDivisión", v: datosFinales.division }, { t: "ccD_encabezado", v: datosFinales.division },
          { t: "ccServicios", v: datosFinales.servicios },
          { t: "ccContrato", v: datosFinales.contrato },
          { t: "ccAPI", v: datosFinales.api },
          { t: "ccProyecto", v: datosFinales.proyecto }, { t: "ccNProyecto_Encabezado", v: datosFinales.proyecto },
          { t: "ccNombreDoc", v: datosFinales.nombreDoc }, { t: "ccNombre doc", v: datosFinales.nombreDoc },
          { t: "ccCodigo", v: datosFinales.codigo },
          { t: "ccRevision", v: datosFinales.revision }
      ];

      for (let item of tagsMapa) {
        if(!item.v) continue;
        let ccs = context.document.contentControls.getByTag(item.t);
        ccs.load("items");
        await context.sync();
        if (ccs.items.length > 0) {
           for (let cc of ccs.items) {
             cc.insertText(item.v, "Replace");
             contadores++;
           }
        }
      }
      
      await context.sync();
      if (msgLabel) msgLabel.textContent = "¡Listo! Datos insertados.";
    });
  } catch (error) {
    console.error(error);
  }
}

// --- 2. HERRAMIENTAS DE FORMATO (CINTA) ---

async function limpiarFormato(event) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      context.load(selection, "font");
      await context.sync();
      selection.font.set({ name: "Arial", size: 11, color: "#000000", bold: false, italic: false });
      await context.sync();
      
      context.load(selection, "paragraphFormat");
      await context.sync();
      try { 
          selection.paragraphFormat.alignment = "Justified"; 
          await context.sync(); 
      } catch (e) { 
          console.warn("No se pudo justificar."); 
      }
    });
  } catch (error) { console.error(error); } 
  finally { if (event) event.completed(); }
}

async function insertarFecha(event) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const fechaHoy = new Date().toLocaleDateString();
    selection.insertText(fechaHoy, "Replace");
    await context.sync();
  });
  if (event) event.completed();
}

// --- 3. ESTILOS FDA (CINTA) ---

async function estiloTitulo1(event) { await aplicarEstiloProfesional("Título 1", "Heading 1"); if (event) event.completed(); }
async function estiloTitulo2(event) { await aplicarEstiloProfesional("Título 2", "Heading 2"); if (event) event.completed(); }
async function estiloTitulo3(event) { await aplicarEstiloProfesional("Título 3", "Heading 3"); if (event) event.completed(); }

async function aplicarEstiloProfesional(nombreEsp, nombreIng) {
  await Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      selection.style = nombreEsp; 
      await context.sync();
    } catch (error) {
      try {
        const selection = context.document.getSelection();
        selection.style = nombreIng;
        await context.sync();
      } catch (e2) {}
    }
  });
}

// --- 4. REGISTRO DE FUNCIONES ---
Office.actions.associate("limpiarFormato", limpiarFormato);
Office.actions.associate("insertarFecha", insertarFecha);
Office.actions.associate("estiloTitulo1", estiloTitulo1);
Office.actions.associate("estiloTitulo2", estiloTitulo2);
Office.actions.associate("estiloTitulo3", estiloTitulo3);