function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp. -> Creamos el menu con submenu
    ui.createMenu('Opciones avanzadas')
        .addItem('Enviar Correo', 'menuItem1')
        .addToUi();
}

function menuItem1() {
    //Selecionamos al hoja SHEET1
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respuestas de formulario 1');

    //Variables formulario
    var asignatura;
    var descripcion;
    var correo = [];
    var fecha;
    var meil = "";
    //Direccion de correo donde se enviará el correo
    var destino = "";
    //Variable que recoge si está el calendario repetido o no
    var coinciden;
    //Cogemos el contenido de la hoja
    var data = ss.getDataRange().getValues();
    var j;
    for (var i = 1; i < data.length; i++) {
        //Comprobamos que no se haya enviado un correo, si el campo tiene de valor 1 es que se ha enviado 
        if (data[i][6] != '1') {
            //Asignatura;
            asignatura = data[i][1];
            //Descripcion de la tarea
            descripcion = data[i][2];
            //Guardamos el correo de destino
            meil = data[i][3];
            correo = meil.split(',');
            //Guardamos la fecha de entrega
            fecha = data[i][4];
            //Comprobamos si la fecha coincide con alguna otra
            coinciden = compruebaFecha(i, 4);
            //Marcamos la fila como leída para que la próxima vez que se ejecute la función no la revise
            ss.getRange(i + 1, 7).setValue('1').setBackground('red');
            if (coinciden) {
                envioMail(asignatura, descripcion, correo, fecha);
            }
            Logger.log('coincide: ' + coinciden);
            //poblarCalendario(asignatura, descripcion, fecha);
        }

    }
}

function compruebaFecha(fila, columna) {
  //Este módulo revisa la fechas de todas las tareas y las compara con la fecha de la tarea recien añadida, en caso
  //de que coincida con la fecha de otra tarea devolverá True
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respuestas de formulario 1');
    var data = ss.getDataRange().getValues();
    var hayCoincidencia = false;
    for (var i = 1; i < data.length; i++) {
        if (i != fila && data[i][columna].toString() == data[fila][columna].toString()) {
            hayCoincidencia = true;
            Logger.log(data[i][columna] + "     " + data[fila][columna]);
        }

    }
    return hayCoincidencia;
}

function envioMail(asignatura, descripcion, correo, fecha) {

    //El mensaje contendrá el cuerpo del correo.
    var mensaje;
    mensaje = "Nombre Asignatura: " + asignatura;
    mensaje = mensaje + "\nDescripción de la tarea: " + descripcion;
    mensaje = mensaje + "\nFecha de entrega: " + fecha;
    mensaje = mensaje + "\nSu tarea coincide con otra/s tareas de otros profesores";

    //armamos el asunto del meil
  var asunto = 'La tarea de ' +asignatura+ ' con fecha ' +fecha+ ' coincide con otra tarea';

    //usamos la API de Gmail para enviar el meil.
    for (var i = 0; i < correo.length; i++) {
        GmailApp.sendEmail(correo[i], asunto, mensaje);

    }
}
