function initProgram() {

  /* creating first sheet with names list */
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //activeSpreadsheet.setActiveSheet(activeSpreadsheet.getSheets()[0]);
  var newSheetName = 'noms';
  var newSheet = activeSpreadsheet.insertSheet(newSheetName, 0);
  //newSheet.setName(newSheetName);

  /* Names list header */
  activeSpreadsheet.getSheetByName(newSheetName).getRange('A1').setValue("Classe");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('B1').setValue("Nom");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('C1').setValue("Commentaire/Comportement");
  /* Names */
  activeSpreadsheet.getSheetByName(newSheetName).getRange('A2').setValue("4A");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('B2').setValue("Jean");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('C2').setValue("Bon comportement");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('A3').setValue("4A");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('B3').setValue("Marie");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('C3').setValue("Bon comportement");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('A4').setValue("4B");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('B4').setValue("Jean");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('C4').setValue("Bon comportement");
  // insert images to assign scripts manually
  insertImageOnSpreadsheet();
}

function insertImageOnSpreadsheet() {
  var SPREADSHEET_URL = getSheetUrl();
  // Name of the specific sheet in the spreadsheet.
  var SHEET_NAME = 'noms';

  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheet = ss.getSheetByName(SHEET_NAME);

  var response = UrlFetchApp.fetch(
      'https://github.com/s3b1/Arlequin/blob/master/assets/new.png?raw=true');
  var binaryData = response.getContent();
  var response2 = UrlFetchApp.fetch(
      'https://github.com/s3b1/Arlequin/blob/master/assets/watch.png?raw=true');
  var binaryData2 = response2.getContent();

  // Insert the image in cell A1.
  var blob = Utilities.newBlob(binaryData, 'image/png', 'NewEval');
  sheet.insertImage(blob, 4, 1);

  var blob2 = Utilities.newBlob(binaryData2, 'image/png', 'WatchProgression');
  sheet.insertImage(blob2, 4, 2);
}

function getSheetUrl() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SS.getActiveSheet();
  var url = '';
  url += SS.getUrl();
  url += '#gid=';
  url += ss.getSheetId();
  return url;
}


function createEval() {

  /* user input name of test */
  var testName = Browser.inputBox("Nom de l'évaluation : ");
  /* user input grade */
  var gradeName = Browser.inputBox("Nom de la classe : ");

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var newSheet = activeSpreadsheet.insertSheet();

  //var lastRow = activeSpreadsheet.getLastRow();
  //var range = activeSpreadsheet.getRange("A2:B" + lastRow);

  /* name of new sheet*/
  var newSheetName = testName + " - " + gradeName;
  newSheet.setName(newSheetName);

  /* list of skills */
  activeSpreadsheet.getSheetByName(newSheetName).getRange('A1').setValue(gradeName);
  activeSpreadsheet.getSheetByName(newSheetName).getRange('B1').setValue("PDS: formuler pb scientifique");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('C1').setValue("PDS: proposer H pr résoudre un pb scientifique");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('D1').setValue("PDS: concevoir exp pr tester H");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('E1').setValue("PDS: ut instruments,techniques de préparation/collectes");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('F1').setValue("PDS: interpréter résultats et tirer cc");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('G1').setValue("PDS: communiquer sur ses démarches,résultats,choix,en argumentant");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('H1').setValue("PDS: id+choisir outils/techniques pr mettre en oeuvre démarche scientifique");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('I1').setValue("CCR: concevoir+mettre en oeuvre protocole expérimentale");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('J1').setValue("UOMA: apprendre à organiser son travail");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('K1').setValue("UOMA: id+choisir outils/techniques pr garder trace de ses recherches");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('L1').setValue("PDL: lire et exploiter des données présentées sous différentes formes");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('M1').setValue("PDL: représenter données sous différentes formes+choisir celle adaptée situation w");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('N1').setValue("UON: conduire recherche information sur internet");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('O1').setValue("UON: ut logiciel d'acquisition de données/simulation et bases de données");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('P1').setValue("ACER: id impacts activités humaines sur l'envt à différentes échelles");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('Q1').setValue("ACER: fonder choix comportement responsable:santé/envt sur arguments scientifiques");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('R1').setValue("ACER: comprendre responsabilités individuelle/collective préservation des ressources");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('S1').setValue("ACER: particper élaboration régles sécurité/appliquer labo/terrain");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('T1').setValue("ACER: distinguer croyance d'une idée et savoir scientifique");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('U1').setValue("SSET: situer l'espèce humaine dans l'évolution des espèces");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('V1').setValue("SSET: appréhender différentes échelles de temps géologique/biologiques");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('W1').setValue("SSET: appréhender différentes échelles/spatiales d'un phénomène/fonction");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('X1').setValue("SSET: id av histoire sciences/techniques cmt se construit savoir scientifique");
  activeSpreadsheet.getSheetByName(newSheetName).getRange('Y1').setValue("Connaissances");

  activeSpreadsheet.getSheetByName(newSheetName).setColumnWidths(1,26, 20);
  activeSpreadsheet.getSheetByName(newSheetName).setColumnWidth(1, 200);
  activeSpreadsheet.getSheetByName(newSheetName).getRange("B1:Y1").setTextRotation(45);

  /* copy last line and change the letter + the label keep capital letter*/

  /* print student names */

  var firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("noms");
  var secondSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName);

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(firstSheet);

  var lastRow = SpreadsheetApp.getActiveSpreadsheet().getLastRow();
  var range = activeSpreadsheet.getRange("A2:B" + lastRow);

  var j = 2;

  for(var i = 1; i < lastRow; i++){

    var iGrade = range.getCell(i, 1).getValue();

     if(iGrade == gradeName){
       var iName = range.getCell(i, 2).getValue();

       SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(secondSheet);
       SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName).getRange("A" + j).setValue(iName);
       //SpreadsheetApp.getUi().alert("eleve num " + j + " " + iName);
       j ++;
       SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(firstSheet);
     }
    }

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(secondSheet);

}

// Use this code for Google Docs, Slides, Forms, or Sheets.
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Dialog')
      .addItem('Open', 'openDialog')
      .addToUi();
}

function openDialog() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeRow = activeSpreadsheet.getActiveRange().getRow();
  var student = activeSpreadsheet.getActiveCell().getValue();
  var range = activeSpreadsheet.getActiveSheet().getRange(activeRow, 1, activeRow, 2);
  var studentClass = range.getCell(1,1).getValue();
  var html = HtmlService.createHtmlOutputFromFile('Index').setWidth(1900).setHeight(900);

  //html.append("<h1>" + student + "</h1>");
  html.append("<table id='myTable' style='display:none;'>");
  html.append("<th>Evaluation</th>");
  html.append("<th>PDS: formuler pb scientifique</th>");
  html.append("<th>PDS: proposer H pr résoudre un pb scientifique</th>");
  html.append("<th>PDS: concevoir exp pr tester H</th>");
  html.append("<th>PDS: ut instruments,techniques de préparation/collectes</th>");
  html.append("<th>PDS: interpréter résultats et tirer cc</th>");
  html.append("<th>PDS: communiquer sur ses démarches,résultats,choix,en argumentant</th>");
  html.append("<th>PDS: id+choisir outils/techniques pr mettre en oeuvre démarche scientifique</th>");
  html.append("<th>CCR: concevoir+mettre en oeuvre protocole expérimentale</th>");
  html.append("<th>UOMA: apprendre à organiser son travail</th>");
  html.append("<th>UOMA: id+choisir outils/techniques pr garder trace de ses recherches</th>");
  html.append("<th>PDL: lire et exploiter des données présentées sous différentes formes</th>");
  html.append("<th>PDL: représenter données sous différentes formes+choisir celle adaptée situation w</th>");
  html.append("<th>UON: conduire recherche information sur internet</th>");
  html.append("<th>UON: ut logiciel d'acquisition de données/simulation et bases de données</th>");
  html.append("<th>ACER: id impacts activités humaines sur l'envt à différentes échelles</th>");
  html.append("<th>ACER: fonder choix comportement responsable:santé/envt sur arguments scientifiques</th>");
  html.append("<th>ACER: comprendre responsabilités individuelle/collective préservation des ressources</th>");
  html.append("<th>ACER: particper élaboration régles sécurité/appliquer labo/terrain</th>");
  html.append("<th>ACER: distinguer croyance d'une idée et savoir scientifique</th>");
  html.append("<th>SSET: situer l'espèce humaine dans l'évolution des espèces</th>");
  html.append("<th>SSET: appréhender différentes échelles de temps géologique/biologiques</th>");
  html.append("<th>SSET: appréhender différentes échelles/spatiales d'un phénomène/fonction</th>");
  html.append("<th>SSET: id av histoire sciences/techniques cmt se construit savoir scientifique</th>");
  html.append("<th>Connaissances</th>");

  /* copy last line and change the label keep <th> tags*/

  html.append(getSkillsValues(student, studentClass));

  html.append('</table>');

  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Progression de ' + student + ' (' + studentClass + ')');
}

function getSkillsValues(pStudent, studentClass) {

  var output = '';

  //loop sheets
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for(var i = 1; i < sheets.length; i++){
    var cell = sheets[i].getRange('A1').getValue();
    //if class matches
    if(cell == studentClass){
      var evalName = sheets[i].getName();
      //get exam name
      var countStudents = sheets[i].getLastRow();
      var countSkills = sheets[i].getLastColumn();
      //loop students
      for(var j = 2; j <= countStudents; j++){

        var rangeStudent = sheets[i].getRange(j, 1, j, countSkills);
        var student = rangeStudent.getCell(1,1).getValue();

        if(student == pStudent){
          output += '<tr><td class="rotate">' + evalName + '</td>';
          //if student matches get skills values
          for(var k = 2; k <= countSkills; k++){

            var skillLabel = sheets[i].getRange(1, k).getValue();
            var skillValue = sheets[i].getRange(j, k).getValue();

            var myClass = '';

            if(skillValue == 1){
             myClass = 'grey';
            } else if (skillValue == 2){
            myClass = 'red';
            } else if (skillValue == 3){
            myClass = 'orange';
            } else if (skillValue == 4){
            myClass = 'green';
            } else {
            myClass = 'empty';
            }

            output += '<td class="' + myClass + '">' + skillValue + '</td>';
          }

        }

        output += '</tr>';

      }

    }

  }
  return output;
}
