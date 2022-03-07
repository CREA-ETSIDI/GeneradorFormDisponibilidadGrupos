function GenerarFormularioEleccionDisponibilidades() {
  let formPrincipal = FormApp.getActiveForm();
  CrearPreguntas(formPrincipal);
  let folder = DriveApp.getFileById(formPrincipal.getId()).getParents().next();
  let censorFolder = folder.createFolder("NO TOCAR");
  DriveApp.getFileById(formPrincipal.getId()).moveTo(censorFolder);
  let hojaRespuestas = DriveApp.getFileById(SpreadsheetApp.create(formPrincipal.getTitle() + "(respuestas)").getId()).moveTo(censorFolder);
  formPrincipal.setDestination(FormApp.DestinationType.SPREADSHEET, hojaRespuestas.getId());
  let semanarioDisponibilidad = SpreadsheetApp.create("Semanario Disponibilidad");
  semanarioDisponibilidad.insertSheet(1).setName("Telemático").deleteColumns(9,18);
  let sheets = semanarioDisponibilidad.getSheets();
  sheets[0].setName("Presencial").deleteColumns(7,20);
  sheets[0].deleteRows(25,1000-25);
  sheets[1].deleteRows(25,1000-25);
  sheets[0].setColumnWidth(1,95);
  sheets[0].setColumnWidths(2,5,245);
  sheets[1].setColumnWidth(1,95);
  sheets[1].setColumnWidths(2,7,245);
  DriveApp.getFileById(semanarioDisponibilidad.getId()).moveTo(folder);
  Generador(semanarioDisponibilidad, hojaRespuestas.getUrl());
}

function CrearPreguntas(formulario){
  clearForm(formulario);

  formulario.setCollectEmail(true);
  formulario.setAllowResponseEdits(true);
  formulario.setLimitOneResponsePerUser(true);

  formulario.addTextItem().setTitle("Introduce tu nombre");

  formulario.addCheckboxGridItem().setTitle("¿En qué franjas horarias puedes reunirte presencialmente? (en la ETSIDI)").setRows(CrearFilas()).setColumns(CrearColumnas().slice(0,5));

  formulario.addCheckboxGridItem().setTitle("¿En qué franjas horarias puedes reunirte virtualmente? (Discord)").setRows(CrearFilas()).setColumns(CrearColumnas());
}

function CrearFilas(){
  let rows = [];
  for(let i = 0; i < 24; i++){
    //Logger.log(String(Math.floor((i/2)+9)) + ":" + (i%2==0?"00":"30"));
    rows.push(String(Math.floor((i/2)+9)) + ":" + (i%2==0?"00":"30")+  "-"  +Math.floor(((i+1)/2)+9) + ":" + ((i+1)%2==0?"00":"30"));
  }
  return rows;
}

function CrearColumnas(){
  return ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
}

//VIVA STACK OVERFLOW JODER
function clearForm(form){
  var items = form.getItems();
  while(items.length > 0){
    form.deleteItem(items.pop());
  }
}

function Generador(SS, myURL) {
  let presencial = SS.getSheets()[0];
  let discordu = SS.getSheets()[1];

  //Formato condicional hoja presencial
  for(let x = 2; x <= 6; x++)
  {
    for(let y = 2; y <= 25; y++)
    {
      presencial.getRange(y,x).setValue(y);
    }
  }

  for(let i = 24; i >= 1; i --)
  {
    let rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=((SUM(B2:F2) >= SUM(B$2:F$2)) + (SUM(B2:F2) >= SUM(B$3:F$3)) + (SUM(B2:F2) >= SUM(B$4:F$4)) + (SUM(B2:F2) >= SUM(B$5:F$5)) + (SUM(B2:F2) >= SUM(B$6:F$6)) + (SUM(B2:F2) >= SUM(B$7:F$7)) + (SUM(B2:F2) >= SUM(B$8:F$8)) + (SUM(B2:F2) >= SUM(B$9:F$9)) + (SUM(B2:F2) >= SUM(B$10:F$10)) + (SUM(B2:F2) >= SUM(B$11:F$11)) + (SUM(B2:F2) >= SUM(B$12:F$12)) + (SUM(B2:F2) >= SUM(B$13:F$13)) + (SUM(B2:F2) >= SUM(B$14:F$14)) + (SUM(B2:F2) >= SUM(B$15:F$15)) + (SUM(B2:F2) >= SUM(B$16:F$16)) + (SUM(B2:F2) >= SUM(B$17:F$17)) + (SUM(B2:F2) >= SUM(B$18:F$18)) + (SUM(B2:F2) >= SUM(B$19:F$19)) + (SUM(B2:F2) >= SUM(B$20:F$20)) + (SUM(B2:F2) >= SUM(B$21:F$21)) + (SUM(B2:F2) >= SUM(B$22:F$22)) + (SUM(B2:F2) >= SUM(B$23:F$23)) + (SUM(B2:F2) >= SUM(B$24:F$24)) + (SUM(B2:F2) >= SUM(B$25:F$25))) = " + i).setBackground(String(presencial.getRange(i+1,2).getBackground())).setRanges([presencial.getRange(2,1,24,1)]).build();
    let rules = presencial.getConditionalFormatRules();
    rules.push(rule);
    presencial.setConditionalFormatRules(rules);
  }

  for(let x = 2; x <= 6; x++)
  {
    for(let y = 2; y <= 25; y++)
    {
      presencial.getRange(y,x).setValue(x);
    }
  }

    for(let i = 5; i >= 1; i --)
  {
    let rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=((SUM(B2:B25) >= SUM(B2:B25)) + (SUM(B2:B25) >= SUM(C2:C25)) + (SUM(B2:B25) >= SUM(D2:D25)) + (SUM(B2:B25) >= SUM(E2:E25)) + (SUM(B2:B25) >= SUM(F2:F25))) = " + i).setBackground(String(presencial.getRange(2,i+1).getBackground())).setRanges([presencial.getRange(1,2,1,5)]).build();
    let rules = presencial.getConditionalFormatRules();
    rules.push(rule);
    presencial.setConditionalFormatRules(rules);
  }


  //Formato condicional hoja telematica

  for(let x = 2; x <= 8; x++)
  {
    for(let y = 2; y <= 25; y++)
    {
      discordu.getRange(y,x).setValue(y);
    }
  }

  for(let i = 24; i >= 1; i --)
  {
    let rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=((SUM(B2:F2) >= SUM(B$2:F$2)) + (SUM(B2:F2) >= SUM(B$3:F$3)) + (SUM(B2:F2) >= SUM(B$4:F$4)) + (SUM(B2:F2) >= SUM(B$5:F$5)) + (SUM(B2:F2) >= SUM(B$6:F$6)) + (SUM(B2:F2) >= SUM(B$7:F$7)) + (SUM(B2:F2) >= SUM(B$8:F$8)) + (SUM(B2:F2) >= SUM(B$9:F$9)) + (SUM(B2:F2) >= SUM(B$10:F$10)) + (SUM(B2:F2) >= SUM(B$11:F$11)) + (SUM(B2:F2) >= SUM(B$12:F$12)) + (SUM(B2:F2) >= SUM(B$13:F$13)) + (SUM(B2:F2) >= SUM(B$14:F$14)) + (SUM(B2:F2) >= SUM(B$15:F$15)) + (SUM(B2:F2) >= SUM(B$16:F$16)) + (SUM(B2:F2) >= SUM(B$17:F$17)) + (SUM(B2:F2) >= SUM(B$18:F$18)) + (SUM(B2:F2) >= SUM(B$19:F$19)) + (SUM(B2:F2) >= SUM(B$20:F$20)) + (SUM(B2:F2) >= SUM(B$21:F$21)) + (SUM(B2:F2) >= SUM(B$22:F$22)) + (SUM(B2:F2) >= SUM(B$23:F$23)) + (SUM(B2:F2) >= SUM(B$24:F$24)) + (SUM(B2:F2) >= SUM(B$25:F$25))) = " + i).setBackground(String(discordu.getRange(i+1,2).getBackground())).setRanges([discordu.getRange(2,1,24,1)]).build();
    let rules = discordu.getConditionalFormatRules();
    rules.push(rule);
    discordu.setConditionalFormatRules(rules);
  }

  for(let x = 2; x <= 8; x++)
  {
    for(let y = 2; y <= 25; y++)
    {
      discordu.getRange(y,x).setValue(x);
    }
  }

    for(let i = 8; i >= 1; i --)
  {
    let rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=((SUM(B2:B25) >= SUM(B2:B25)) + (SUM(B2:B25) >= SUM(C2:C25)) + (SUM(B2:B25) >= SUM(D2:D25)) + (SUM(B2:B25) >= SUM(E2:E25)) + (SUM(B2:B25) >= SUM(F2:F25)) + (SUM(B2:B25) >= SUM(G2:G25)) + (SUM(B2:B25) >= SUM(H2:H25))) = " + i).setBackground(String(discordu.getRange(2,i+1).getBackground())).setRanges([discordu.getRange(1,2,1,7)]).build();
    let rules = discordu.getConditionalFormatRules();
    rules.push(rule);
    discordu.setConditionalFormatRules(rules);
  }

  let dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];

  //Generador de funciones hoja presencial
  for (let x = 0; x < 5; x++)
  {
    for(let y = 2; y <= 25; y++)
    {
      let columnas = columnToLetter(y+2) + ':' + columnToLetter(y+2);
      presencial.getRange(y, x+2).setValue('=COUNTIF(IMPORTRANGE("' + myURL + '";"Form Responses 1!'+columnas+'");"*'+dias[x]+'*")');
    }
  }

  //Generador de funciones hoja telematica
  for (let x = 0; x < 7; x++)
  {
    for(let y = 2; y <= 25; y++)
    {
      let columnas = columnToLetter(y+26) + ':' + columnToLetter(y+26);
      discordu.getRange(y, x+2).setValue('=COUNTIF(IMPORTRANGE("' + myURL + '";"Form Responses 1!'+columnas+'");"*'+dias[x]+'*")');
    }
  }

  presencial.getRange(1,2,1,5).setValues([CrearColumnas().slice(0,5)]);
  discordu.getRange(1,2,1,7).setValues([CrearColumnas()]);

  let filas = CrearFilas();
  for(let y = 0; y < 24; y++){
    presencial.getRange(y+2,1).setValue(filas[y]);
    discordu.getRange(y+2,1).setValue(filas[y]);
  }
  presencial.getRange(1, 1).setValue('=IMPORTRANGE("' + myURL + '";"Form Responses 1!D:D")');
}

//VIVA STACK OVERFLOW JODER
function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
