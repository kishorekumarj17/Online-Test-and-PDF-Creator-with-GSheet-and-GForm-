// Google Sheet Menu Creation
function onOpen(e) 
{
  SpreadsheetApp.getUi() 
      .createMenu('Online Test Menu')
      .addItem('Online Test Template', 'testTemplate')
      .addItem('Create test by selecting range', 'createTest')
      .addToUi();
}

// template Generator
function testTemplate()
{
  var lastRow=SpreadsheetApp.getActiveSheet().getLastRow();
  SpreadsheetApp.getActiveSheet().appendRow(["Test Name:"]);
  SpreadsheetApp.getActiveSheet().appendRow(["Question","Option 1","Option 2","Option 3","Option 4","Correct Option"]);
  SpreadsheetApp.getActiveSheet().insertRowBefore(lastRow+1);
  SpreadsheetApp.getActiveSheet().getRange("A"+(lastRow+2)+":F"+(lastRow+3)).setBackground("red").setFontColor("white");
  SpreadsheetApp.getActiveSheet().getRange(lastRow+2,2).setBackground("yellow").setFontColor("black");
  SpreadsheetApp.getActiveSheet().setCurrentCell(SpreadsheetApp.getActiveSheet().getRange("A"+(lastRow+1+1+2)))
  SpreadsheetApp.getUi().alert("Successfully created test template. Enter questions, options and answers.");
}

// test creater function
function createTest()
{
  var qrangev=SpreadsheetApp.getActiveSheet().getActiveRange();
  var qrange=qrangev.getValues();
  var questions=[];
  console.log(qrange[0][0]=="Test Name:")
  if(qrange.length>2 && qrange[0].length==6 && qrange[0][0]=="Test Name:")
  {
        if(qrange[0][1].trim()=="")
          {
            SpreadsheetApp.getUi().alert("Missing Details. Please enter test name.");
          }
        else
        {
          for(let i=2;i<qrange.length;i++)
          {
            if(qrange[i][0].length!=0&&qrange[i][1].length!=0&&qrange[i][2].length!=0&&qrange[i][3].length!=0&&qrange[i][4].length!=0&&qrange[i][5]. length!=0)
            {
              if(typeof(parseInt(qrange[i][5])=="number") && parseInt(qrange[i][5]).toString()!="NaN")
              {
                questions.push(qrange[i])
              }
              else
              {
                SpreadsheetApp.getUi().alert("Please enter 'Correct option' as number from 1 to 4");
                break;
              }
            }
            else
            {
              SpreadsheetApp.getUi().alert("Missing Details. Please check questions are filled correctly.");
              break;
            }
          }
        }
  }
  else
  {

    SpreadsheetApp.getUi().alert("Missing Details. Please select a range with questions and test name.");

  }
  if(questions.length==qrange.length-2)
  {

    var form=FormApp.create(qrange[0][1].trim());
    var doc=DocumentApp.create(qrange[0][1].trim());
    doc.getBody().appendParagraph(qrange[0][1].trim()).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setLineSpacing(2).editAsText().setFontFamily("Times New Roman").setFontSize(16).setBold(true);

    form.setIsQuiz(true);
    form.setProgressBar(true)
    for(let i=2;i<qrange.length;i++)
    {

       var item=form.addMultipleChoiceItem();

      doc.getBody().appendParagraph(qrange[i][0].trim()).setLineSpacing(1.5).editAsText().setFontFamily("Times New Roman").setFontSize(14);
       item.setTitle(qrange[i][0].trim())
          .setChoices([
        item.createChoice(qrange[i][1],1==parseInt(qrange[i][5])),
        item.createChoice(qrange[i][2],2==parseInt(qrange[i][5])),
        item.createChoice(qrange[i][3],3==parseInt(qrange[i][5])),
        item.createChoice(qrange[i][4],4==parseInt(qrange[i][5]))
        ])
          .setRequired(true)
          .setPoints(1);

        doc.getBody().appendParagraph("a. "+qrange[i][1]).setLineSpacing(1.15).editAsText().setFontFamily("Times New Roman").setFontSize(12).setBold(1==parseInt(qrange[i][5])).setItalic(1==parseInt(qrange[i][5]));
        doc.getBody().appendParagraph("b. "+qrange[i][2]).setLineSpacing(1.15).editAsText().setFontFamily("Times New Roman").setFontSize(12).setBold(2==parseInt(qrange[i][5])).setItalic(2==parseInt(qrange[i][5]));
        doc.getBody().appendParagraph("c. "+qrange[i][3]).setLineSpacing(1.15).editAsText().setFontFamily("Times New Roman").setFontSize(12).setBold(3==parseInt(qrange[i][5])).setItalic(3==parseInt(qrange[i][5]));
        doc.getBody().appendParagraph("d. "+qrange[i][4]).setLineSpacing(1.15).editAsText().setFontFamily("Times New Roman").setFontSize(12).setBold(4==parseInt(qrange[i][5])).setItalic(4==parseInt(qrange[i][5]));
    console.log(qrange[i][3])
    var answerop=(1==parseInt(qrange[i][5]))? qrange[i][1] : (2==parseInt(qrange[i][5]) ? qrange[i][2]:(3==parseInt(qrange[i][5])?qrange[i][3]:(4==parseInt(qrange[i][5])?qrange[i][4]:'No option')))
        doc.getBody().appendParagraph("Answer : "+answerop).setLineSpacing(2).editAsText().setFontFamily("Times New Roman").setFontSize("12");
    }
    doc.saveAndClose();
    form.setShowLinkToRespondAgain(false);
    form.setLimitOneResponsePerUser(true);

    var blob=doc.getAs("application/pdf");
    blob.setName(qrange[0][1].trim()+".pdf")
    var pdffile=DriveApp.createFile(blob);
    pdffile.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    var labels = { 'labels': {restricted: true} }; 
    Drive.Files.update(labels,pdffile.getId())


    form.setConfirmationMessage("Test submitted successfully. Please click below button to see the score. PDF File: "+pdffile.getUrl())
    if(SpreadsheetApp.getActiveSheet().getRange("C"+(parseInt(qrangev.getRow()))).setValue(form.shortenFormUrl(form.getPublishedUrl())))
    {
      if(SpreadsheetApp.getActiveSheet().getRange("E"+(parseInt(qrangev.getRow()))).setValue(pdffile.getUrl()))
      {
          SpreadsheetApp.getUi().alert("Test created successfully.");
      }
    }

  }
}

