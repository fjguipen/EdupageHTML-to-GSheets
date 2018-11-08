/**************************************************************

  Funcion que obtiene todos los datos del archivo HTML y los 
  guarda en una variable. Tambien damos el nombre de la hoja
  de Excel a crear para volcar todos los datos.

**************************************************************/
function recuperarDatos(datos) {
  var nombreExcel = "Asistencias NP";
  /** Llamamos a la funcion crearExcel pasandole el nombre de la hoja
   y el html con todos los datos. **/
  crearExcel(nombreExcel,datos);
}

/**********************************************************************

  Funcion que crea la nueva hoja de Excel con el nombre que le hemos
  dado y crea en esa hoja el encabezado con los titulos en la primera
  fila de la hoja.

**********************************************************************/
function crearExcel(nombre,html) {
  var nuevaHoja = SpreadsheetApp.getActive();
  var sheet = nuevaHoja.getActiveSheet();
  sheet.clear();
  sheet.getRange(1,1).setValue('CURSO').setBackgroundRGB(188, 230, 247);
  sheet.getRange(1,2).setValue('CICLO').setBackgroundRGB(188, 230, 247);
  sheet.getRange(1,3).setValue('ALUMNO').setBackgroundRGB(188, 230, 247);
  sheet.getRange(1,4).setValue('ASIGNATURA').setBackgroundRGB(188, 230, 247);
  sheet.getRange(1,5).setValue('USUARIO').setBackgroundRGB(188, 230, 247);
  sheet.getRange(1,6).setValue('PRESENTE').setBackgroundRGB(188, 230, 247);
  sheet.getRange(1,7).setValue('AUSENTE').setBackgroundRGB(188, 230, 247);
  sheet.getRange(1,8).setValue('%').setBackgroundRGB(188, 230, 247);
  
  introducirDatos(sheet,html);
}

/****************************************************

  

****************************************************/
function introducirDatos(nombreDoc,datos) { 
  var matriz = [];

  //Introducir todos los datos en una matriz.
  datos.forEach(function (data){
    data.alumnos.forEach(function (alumno) {    
      alumno.asignaturas.forEach(function (asignatura) {
        var fila = [];
        fila[0] = data.curso.substr(0,2);
        fila[1] = data.curso.substr(2);
        fila[2] = alumno.nombre;
        fila[3] = asignatura.id;
        fila[4] = "";
        fila[5] = asignatura.asistencias;
        fila[6] = asignatura.faltas;
        if (!asignatura.porcentaje == "") {
          fila[7] = fila[7] = asignatura.porcentaje+"%";
        }else {
          fila[7] = "";
        }
        
        matriz.push(fila);
      })    
    })
  })
  
  //Escribir en el excel todos los datos de la matriz
  nombreDoc.getRange(2, 1, matriz.length, matriz[0].length).setValues(matriz);
}


















