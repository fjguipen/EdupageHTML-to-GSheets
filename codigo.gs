/************************************************************************
  Funcion que llama directamente a la funcion onOpen cuando se instala.
************************************************************************/
function onInstall(e) {
  onOpen();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()  //Add a new option in the Google Docs Add-ons Menu
    .addItem("Obtener datos de alumno", "main")
    .addToUi();  //Run the showSidebar function when someone clicks the menu
}

/**
 * Function principal desde la que obtenemos el fichero .html
 * y generamos un string con el contenido de este.
 * 
 * Devuelve un objeto de JS
 */

function main() {
    //Declaracion de variables
    var fileURL = Browser.inputBox('URL del archivo en Drive')
    var id = getParamFromURL(fileURL, "id");
    if (!id){
      Browser.msgBox("URL no válida")
      return
    }
 
    try{
      //Capturamos el error que lanza .getFileById() si no encuentra el archivo
      var file = DriveApp.getFileById(id);
      var rawData;
      var HTMLData;
      var splitedHTML;
      var JSONData = [];
      //Extraemos su contenido y lo parseamos a XML
      rawData = file.getBlob().getDataAsString()
      HTMLData = parsearHTML(rawData);
      //Dividimos el contenido en diferentes páginas para cada tabla
      splitedHTML = getHTMLPages(HTMLData)
      //Para cada página, obtenemos un objeto tipo JSON 
      for (var i = 0; i < splitedHTML.length; i++) {
        var data = toJSON(splitedHTML[i])
        if (data) {
          JSONData.push(data)
        }
      }
      recuperarDatos(JSONData);
      //Avisamos del fin de programa
      Browser.msgBox("Datos recuperados")
    }catch(err){
      Browser.msgBox("El archivo no existe");
    }
  
}

/**
 * Función que acepta un elemento parseado de tipo HTML/XML y
 * obtiene la informacion de los alumnos a aparir de el
 */

function toJSON(rootElement) {
    var data = {
        curso: "",
        fecha: { desde: "", hasta: "" },
        alumnos: []
    }
    var heading;
    var encabezados;
    //Obtenemos los elementos dentro del <div/> que contiene el curso y la fecha
    heading = getElementsByClassName(rootElement, "print-header")[0].getChildren()
    data.curso = getCurso(heading[0].getText());
    data.fecha.desde = heading[1].getText().split("-")[0].trim();
    data.fecha.hasta = heading[1].getText().split("-")[1].trim();
    //Obtenemos los nombres de las columnas (asignaturas)
    encabezados = getEncabezados(rootElement)
    //Control de tablas vacías
    if (!encabezados) {
        return
    }

    /**
     * Obtenemos todos los nodos hijos  de <tbody/> (son los <tr/> que constiuyen a cada alumno)
     * Para cada uno, extraemos la informacion en ellos y la cargamos en nuestro objeto
     * */

    getElementsByTagName(rootElement, "tbody")[0].getChildren().forEach(function (rawAlumno) {
        data.alumnos.push(getInfoFromAlumn(rawAlumno, encabezados))
    })

    return data;
}

/**
 * Esta funcion recibe un string html y lo parsea a XML,
 * eliminando en primer lugar el contenido conflictivo
 * del string.
 */

function parsearHTML(rawData) {
    var cleanData;
    var html;
    //Pasamos por parámetro las etiquetas que queremos eliminar
    cleanData = cleanHTMLTags(rawData, [/<br>/g, /<colgroup(.)*>(.)*<\/colgroup>/g]);
    //Una vez limpio, parseamos
    html = XmlService.parse(cleanData);

    return html;
}

/**
 * Funcion que recibe un string HTML/XML  y un array de expresiones regulares
 * Elimina el contenido que coincida con las expresiones regulares pasadas
 */

function cleanHTMLTags(rawData, regexTags) {
    var cleanData;
    //Valor por defecto, si no se introduce ningun valor -> array vacío
    regexTags = regexTags || []
    //Si se introduce un unico elemento y no es un array, lo convertimos en uno
    if (!Array.isArray(regexTags)){
        regexTags = [regexTags];
    }
    //Obtenemos el contenido dentro de <body/>
    cleanData = rawData.split("<body")[1].split("</body>")[0]
    //Añado de nuevo la apertura y cierre para obtener un elemento completo
    cleanData = '<body' + cleanData + '</body>';
    //Ejecutamos la limpeiza para todas las exp. regulares
    regexTags.forEach(function (e) {
        cleanData = cleanData.replace(e, "")
    })

    return cleanData;
}

//Divide el contenido del HTML en páginas, en funcion del <div/> de separación
function getHTMLPages(html) {
    var root = getElementsByClassName(html, "print-sheet")[0]
    var pages;
    pages = root.getChildren().filter(function (e) {
        return e.asElement() && !e.getAttribute("class")
    })

    return pages
}

//Funcion que obtiene el curso a partir de una cadena de comparación  -> "Ausente, "
function getCurso(str) {
    str = str.split("-");
    str = str[str.length - 1].replace("Ausente, ", "").trim()

    return str
}

/**
 * Funcion que captura los <span/> contenidos en los <th/> dentro del <thead/>,
 * y devuelve su contenido (valores de la primera fila de la tabla)  
 */

function getEncabezados(root) {
    var encabezados = [];

    getElementsByTagName(getElementsByTagName(root, "thead")[0], "th").forEach(function (e) {
        var text = getElementsByTagName(e, "span")[0].getText();
        encabezados.push(text);

    })
    //Devuelvo eliminando el primer elemento ("Estudiantes")  
    return encabezados.slice(1);
}

/**
 * Función que captura la información del alumno <tr/> pasado por parámetro.
 * Para ello, obteine un array con los <td/>, valida si tiene o no el atributo
 * title, el cual porta la informacion sobre la asistencia, y carga toda esa i
 * informacion en un objeto de JS. 
 */

function getInfoFromAlumn(rawAlumno, encabezados) {
    var datosDeAlumno = rawAlumno.getChildren();
    //Logger.log(datosDeAlumno)
    var alumno = {
        //Obtengo el nombre y a su vez lo elimino del array
        nombre: datosDeAlumno.shift().getText(),
        asignaturas: []
    }
    //Para cada <td/>, extraemos la informacion que contienen
    for (var i = 0; i < datosDeAlumno.length; i++) {
        var asignatura = {
            /* Nos valemos del array encabezados para determinar la asignatura,
            aprovechando que existe un orden y los índices coinciden. */
            id: encabezados[i],
            asistencias: "",
            faltas: "",
            porcentaje: ""
        }
        var ausencia = datosDeAlumno[i].getAttribute("title")
        //Nos traemos el atributo "title", si existe, entonces capturamos su contenido
        if (ausencia) {
            //Buscamos un la cadena "Presente: X" (siendo X cualquier digito) y obtenemos el valor
            asignatura.asistencias = ausencia.getValue().match(/Presente: \d+/)[0].split(":")[1].trim()
            //Buscamos un la cadena "Ausente: X" (siendo X cualquier digito) y obtenemos el valor
            asignatura.faltas = ausencia.getValue().match(/Ausente: \d+/)[0].split(":")[1].trim()
            //Calculamos el pordentaje
            asignatura.porcentaje = Math.round(asignatura.faltas / (Number(asignatura.asistencias) +
                                    Number(asignatura.faltas)) * 100);
        }
        alumno.asignaturas.push(asignatura)
    }
    return alumno
}

/***********************************
*                                  *
*       Herramientas y trabajo     *
*       com XML                    *
*                                  *
************************************/

/**
* Funcion que recibe una url y un parametro,
* devolviendote el valor de ese parámetro si lo encuentra en la url
* Devuelve undefined si no encuentra el parametro.
*/ 

function getParamFromURL(url, param){
  var paramsMap = [];
  var paramRetorno;
  try{
    url.split("?")[1].split("&").forEach(function (e){
      var paramStr = e.split("=")
      paramsMap.push({name:paramStr[0], value:paramStr[1]})
    })
    
    paramsMap.forEach(function (e){
      if (e.name === param){
        paramRetorno = e.value
      }
    })
  }catch(err){
    //Logger.log(err.message);
  }
  
  return paramRetorno;
}

/**
 * Funcion que acepta un ID y devuelve el primer elemento que encuentre
 * dentro de root con ese ID
 */

function getElementById(root, idToFind) {
    var element;

    root.getDescendants().forEach(function (e) {
        //Element.asElement() valida que sea un elemento, en caso contrario devuelve null
        if (e.asElement() && e.getAttribute("id").getValue() === idToFind) {
            element = e;
        }
    })

    return element;
}

/**
 * Funcion que acepta un nombre de clase, y devuelve un arary con
 * todo los elementos que encuentre con dicha clase
 */

function getElementsByClassName(root, classNameToFind) {
    var elements = [];

    root.getDescendants().forEach(function (e) {
        //Element.asElement() valida que sea un elemento, en caso contrario devuelve null
        if (e.asElement()) {
            var classes = e.getAttribute("class")
            if (classes && classes.getValue().split(" ").indexOf(classNameToFind) >= 0) {
                elements.push(e);
            }
        }
    })
    return elements;
}

/**
 * Funcion que acepta un nombre de etiqueta, y devuelve un arary con
 * todo los elementos que encuentre con esa etiqueta
 */

function getElementsByTagName(root, tagNameToFind) {
    var elements = [];

    root.getDescendants().forEach(function (e) {
        //Element.asElement() valida que sea un elemento, en caso contrario devuelve null
        if (e.asElement() && e.getName() === tagNameToFind) {
            elements.push(e)
        }
    })
    return elements;
}