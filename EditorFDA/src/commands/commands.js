/* commands.js */

// 1. Inicializar Office. Esto le dice a Word "Estoy listo y escuchando"
Office.onReady(() => {
  // Si necesitas hacer algo al iniciar, va aquí.
  console.log("Commands.js cargado y listo");
});

// 2. Tu función
function insertarFecha(event) {
  Word.run(function (context) {
    var body = context.document.body;
    var hoy = new Date().toLocaleDateString();
    
    // Insertamos al inicio
    body.insertParagraph("FECHA DESDE COMMANDS: " + hoy, "Start");
    
    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
  })
  .finally(function () {
      // 3. ¡CRÍTICO! Avisar a Word que terminamos.
      // Si no pones esto, el botón se queda "pegado" y no funciona la segunda vez.
      if (event) {
          event.completed();
      }
  });
}

// 4. Mapeo (Vinculación)
// El primer nombre 'insertarFecha' debe coincidir con <FunctionName> en el Manifest.
// El segundo nombre es la función de arriba.
Office.actions.associate("insertarFecha", insertarFecha);