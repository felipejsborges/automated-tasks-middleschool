function AdicionarColunaNaPlanilhaPorMatéria() {
  // dados gerais
  var week = '08/09 a 11/09';
  var weekCounter = 2; 
  var subjectName = 'Ed Física - Anos Iniciais';
  
  // dados das planilhas semanais padrão
  var weekFolderId = 'fakeWeekFolderId';
  var subjectColumnLetter = 'K';  
  
  // dados da planilha separada por matéria
  var subjectFolder = 'fakeSubjectFolder';
  var subjectWeekColumnToAdd = 6;
  var subjectWeekColumnToAddLetter = 'F';
  
  AddColumOnTeacherSheets(week, weekCounter, subjectName, weekFolderId, subjectColumnLetter, subjectFolder, subjectWeekColumnToAdd, subjectWeekColumnToAddLetter)
}

function AlimentarBanco08a11setembro(){  
  // id do banco de dados
  var databaseId = "fakeDatabaseID";
  
  // id da pasta do forms da semana
  var weekFormsFolderId = "fakeWeekFormsFolderID";
  
  fillWeekDatabase(databaseId, weekFormsFolderId);  
}

function AlimentarBanco31agostoA04setembro(){  
  // id do banco de dados
  var databaseId = "fakeDatabaseID";
  
  // id da pasta do forms da semana
  var weekFormsFolderId = "fakeWeekFormsFolderID";
  
  fillWeekDatabase(databaseId, weekFormsFolderId);  
}

function AlimentarBanco24a28agosto(){  
  // id do banco de dados
  var databaseId = "fakeDatabaseID";
  
  // id da pasta do forms da semana
  var weekFormsFolderId = "fakeWeekFormsFolderID";
  
  fillWeekDatabase(databaseId, weekFormsFolderId);  
}

function AlimentarResumoDaSemana() {  
  // id da planilha de resumo da semana
  var overviewSheetId = "fakeOverviewSheetID";
  
  // id da pasta com as planilhas de monitoramento da semana
  var monitWeekFolderID = "fakeMonitWeekFolderID";
  
  // semana escrita por extenso
  var week = "Semana de 08-09 a 11-09";
  
  fillOverviewSheet(overviewSheetId, monitWeekFolderID, week);  
}

function ApagarInformacoesDoBancoDeDados() {  
  // id do banco de dados para limpar
  var databaseId = "fakeDatabaseID";
  
  clearDb(databaseId);
}

function fillWeekDatabase(databaseId, weekFormsFolderId) {  
  var idList = [];
  while (aIaF_Folders.hasNext()){
    var id = aIaF_Folders.next().getId();
    idList.push(id)
  }
  
  var formsResponses = []  
  var formsResponsesAF = getFormsResponses(idList[0])  
  var formsResponsesAI = getFormsResponses(idList[1])
  formsResponses = [...formsResponsesAF, ...formsResponsesAI]
  
  clearDb(databaseId)
  
  fillDatabase(formsResponses, databaseId) 
}

function getFormsResponses(folderId, week) {
  var gradeFoldersList = DriveApp.getFolderById(folderId).getFolders();  
  
  var responses = []
  
  while(gradeFoldersList.hasNext()) {
    var gradeFolder = gradeFoldersList.next();
    var gradeFolderId = gradeFolder.getId();
    var gradeFolderName = gradeFolder.getName();
    
    if (gradeFolderName.slice(0, 5).toLowerCase().trim() == "canto" || gradeFolderName.slice(0, 6).toLowerCase().trim() == "escola") {
      continue;
    }
    
    var formfiles = DriveApp.getFolderById(gradeFolderId).getFilesByType(MimeType.GOOGLE_FORMS)    
    
    while(formfiles.hasNext()) {
      var formFile = formfiles.next()
      var formId = formFile.getId();
      var materia = formFile.getName();
      
      var form = FormApp.openById(formId);
      var formResponses = form.getResponses();
      
      for (var i = 0; i < formResponses.length; i++) {
        var itemResponses = formResponses[i].getItemResponses();
        var gradeClass = itemResponses[0].getResponse();
        var studentName = itemResponses[1].getResponse();
        
        var parsedGrade = gradeClass.slice(0, 1) + gradeClass.slice(gradeClass.length - 1, gradeClass.length)
        
        var responseData = {};
        responseData['nome_do_aluno'] = studentName;
        responseData['materia'] = materia;
        responseData['turma'] = parsedGrade;        
        
        responses.push(responseData);
      }      
    }  
  }
  return responses;
}

function fillDatabase(formData, databaseId) {
  var db = SpreadsheetApp.openById(databaseId);
  var sheet = db.getSheets()[0];
  
  var startLine = 1;
  
  while(!sheet.getRange("A:A").getCell(startLine, 1).isBlank()) {
    startLine++;
  }
  
  formData.forEach(function(item, index){
    var values = [
      [item.nome_do_aluno.toLowerCase(), item.materia.toLowerCase(), item.turma.toLowerCase()]
    ];
    var line = startLine + index
    var dataRange = sheet.getRange(line, 1, 1, 3);
    
    dataRange.setValues(values);
  })  
}

function clearDb(dbId) {
  var db = SpreadsheetApp.openById(dbId);
  var sheet = db.getSheets()[0];
  
  var rangeToDelete = sheet.getRange("2:10000");
  rangeToDelete.clear();
}

function fillOverviewSheet(overviewSheetId, monitWeekFolderID, week) {
  var overviewSheet = SpreadsheetApp.openById(overviewSheetId);
  var sheet = overviewSheet.getSheets()[0];  
  sheet.getRange("B3").setValue(week);
 
  var monitGradeFiles = DriveApp.getFolderById(monitWeekFolderID).getFiles()  
  while(monitGradeFiles.hasNext()) {
    var monitGradeFile = monitGradeFiles.next();
    var gradeId = monitGradeFile.getId();
    var grade = monitGradeFile.getName();
    var parsedGrade = grade.trim().slice(0,1) + grade.trim().slice(grade.trim().length - 1, grade.trim().length);
    
    parsedGrade == "1A" && sheet.getRange(100, 2).setValues([[gradeId]]);
    parsedGrade == "1B" && sheet.getRange(100, 3).setValues([[gradeId]]);
    parsedGrade == "1C" && sheet.getRange(100, 4).setValues([[gradeId]]);
    parsedGrade == "2A" && sheet.getRange(100, 5).setValues([[gradeId]]);
    parsedGrade == "2B" && sheet.getRange(100, 6).setValues([[gradeId]]);
    parsedGrade == "2C" && sheet.getRange(100, 7).setValues([[gradeId]]);
    parsedGrade == "3A" && sheet.getRange(100, 8).setValues([[gradeId]]);
    parsedGrade == "3B" && sheet.getRange(100, 9).setValues([[gradeId]]);
    parsedGrade == "3C" && sheet.getRange(100, 10).setValues([[gradeId]]);
    parsedGrade == "4A" && sheet.getRange(100, 11).setValues([[gradeId]]);
    parsedGrade == "4B" && sheet.getRange(100, 12).setValues([[gradeId]]);
    parsedGrade == "4C" && sheet.getRange(100, 13).setValues([[gradeId]]);
    parsedGrade == "5A" && sheet.getRange(100, 14).setValues([[gradeId]]);
    parsedGrade == "5B" && sheet.getRange(100, 15).setValues([[gradeId]]);
    parsedGrade == "5C" && sheet.getRange(100, 16).setValues([[gradeId]]);
    parsedGrade == "6A" && sheet.getRange(100, 17).setValues([[gradeId]]);
    parsedGrade == "6B" && sheet.getRange(100, 18).setValues([[gradeId]]);
    parsedGrade == "6C" && sheet.getRange(100, 19).setValues([[gradeId]]);
    parsedGrade == "7A" && sheet.getRange(100, 20).setValues([[gradeId]]);
    parsedGrade == "7B" && sheet.getRange(100, 21).setValues([[gradeId]]);
    parsedGrade == "7C" && sheet.getRange(100, 22).setValues([[gradeId]]);
    parsedGrade == "8A" && sheet.getRange(100, 23).setValues([[gradeId]]);
    parsedGrade == "8B" && sheet.getRange(100, 24).setValues([[gradeId]]);
    parsedGrade == "8C" && sheet.getRange(100, 25).setValues([[gradeId]]);
    parsedGrade == "9A" && sheet.getRange(100, 26).setValues([[gradeId]]);
    parsedGrade == "9B" && sheet.getRange(100, 27).setValues([[gradeId]]);
    parsedGrade == "9C" && sheet.getRange(100, 28).setValues([[gradeId]]);
  }
}


function AddColumOnTeacherSheets(week, weekCounter, subjectName, weekFolderId, subjectColumnLetter, subjectFolder, subjectWeekColumnToAdd, subjectWeekColumnToAddLetter) {  
  var gradeFilesFromWeek = DriveApp.getFolderById(weekFolderId).getFiles();
  var teacherFiles = DriveApp.getFolderById(subjectFolder).getFiles();
  
  var gradesIdsAndNames = [];  
  while (gradeFilesFromWeek.hasNext()){
    var tmp = {};
    var file = gradeFilesFromWeek.next();
    
    tmp['id'] = file.getId();    
    var name = file.getName();    
    tmp['name'] = name.trim().slice(0,1) + name.trim().slice(name.trim().length - 1, name.trim().length);
    
    gradesIdsAndNames.push(tmp);
  }
  
  while (teacherFiles.hasNext()){
    var file = teacherFiles.next();
    
    var id = file.getId();    
    var unParsedName = file.getName();    
    var name = unParsedName.trim().slice(0,1) + unParsedName.trim().slice(unParsedName.trim().length - 1, unParsedName.trim().length);        
   
    var spreadsheet = SpreadsheetApp.openById(id);
    var sheet = spreadsheet.getSheets()[0];
    
    sheet.getRange(5, 1).activate();
    sheet.getCurrentCell().setValue(subjectName);
    
    sheet.getRange(53, 3).activate();
    sheet.getCurrentCell().setValue(weekCounter);
    
    sheet.insertColumns(subjectWeekColumnToAdd);
    
    sheet.getRange(7, subjectWeekColumnToAdd).activate();
    sheet.getCurrentCell().setValue(week);
    
    sheet.setColumnWidth(subjectWeekColumnToAdd, 80);
    
    sheet.getRange(7, subjectWeekColumnToAdd, 2).activate()
    .mergeVertically()
    .setWrap(true);
    
    sheet.getRange(subjectWeekColumnToAddLetter + '7:' + subjectWeekColumnToAddLetter + '48').activate();
    sheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    
    for (var i = 0; i < gradesIdsAndNames.length; i++) {
      if (gradesIdsAndNames[i].name === name) {
        sheet.getRange(100, subjectWeekColumnToAdd).setValues([[gradesIdsAndNames[i].id + '?now()']]);        
        sheet.getRange(100, subjectWeekColumnToAdd).activate();
        sheet.getActiveRangeList().setFontColor('#ffffff');
        break;
      }
    }
    
    for (var i = 9; i <=48; i++) {
      var query = "=IMPORTRANGE($" + subjectWeekColumnToAddLetter + "$100; \"Plan1!" + subjectColumnLetter + i + "\")";
      sheet.getRange(i, subjectWeekColumnToAdd).setValue(query);
      
      sheet.getRange(i, 4).setValue("=COUNTIF(E" + i + ":" + subjectWeekColumnToAddLetter + i + ";\"X\")");      
    }
  }
}
