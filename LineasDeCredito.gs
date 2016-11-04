function enviarEmail(){
  
  //Declaracion de variables
  var nombre;
  var apellidos;
  var email;
  var dir;
  var prov;
  var cp;
  var ingresos;
  var gastos;
  var esTrabajador;
  var estadoCivil;
  var fecha;
  var hoy;
  var anio;
  var mes;
  var dia;
  
  //Obtencion hoja de calculo
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var row=sheet.getLastRow();
  
  //Obtencion de valores para cada variable
  nombre = sheet.getRange("B"+row).getValue();
  apellidos = sheet.getRange("C"+row).getValue();
  email = sheet.getRange("D"+row).getValue();
  dir = sheet.getRange("E"+row).getValue();
  prov = sheet.getRange("F"+row).getValue();
  cp = sheet.getRange("G"+row).getValue();
  estadoCivil = sheet.getRange("H"+row).getValue();
  if(estadoCivil=="Casado"){
    ingresos = sheet.getRange("I"+row).getValue();
    gastos = sheet.getRange("J"+row).getValue();
    esTrabajador = sheet.getRange("K"+row).getValue();
  }else{
    ingresos = sheet.getRange("L"+row).getValue();
    gastos = sheet.getRange("M"+row).getValue();  
    esTrabajador = sheet.getRange("N"+row).getValue();
  }
  hoy = new Date();
  mes = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "MM");
  dia = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd");
  anio = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "YYYY");
  fecha = dia + "/" + mes + "/" + anio;
  
  //Evaluacion credito
  var credito = "denegada";
  if((esTrabajador=="Si"&&ingresos>gastos)||(estadoCivil=="Casado"&&esTrabajador=="Si"&&ingresos>=2*gastos)){
    credito = "aceptada";
  }  
  
  //Elaboracion E-mail
  var subject = "Resolucion de Solicitud de Linea de Credito";
  var cabecera = "\n\n" + nombre + " " + apellidos  + "\n" + dir + "\n" + prov + "\t" + cp + "\n\n\n";
  var body = fecha + cabecera;
  body = body + "Estimado " + nombre + " con esta carta se le hace de su conocimiento de que su línea de credito ha sido " + credito;
  body = body + "\n\n";
  if(credito == "aceptada"){
    body = body + "En breve nos pondremos en contacto con usted.\n\nUn cordial saludo."
  }else{
    body = body + "Sin más por el momento nos despedimos."
  }  
  Logger.log(body);
  GmailApp.sendEmail(email, subject, body); 
}

  
