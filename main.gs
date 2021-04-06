/**
 * Created with generous assistance from Michelle Suranofsky, Lehigh University, and other helpful folk on the Interwebs. 
 * Thank you all for helping me with this journey!!
 * 2021-03-25
 * 
 * Need to figure out how to report if holdings not set for AS0 for an OCLC number!!
 */

function runReport(form) {

  var ui = SpreadsheetApp.getUi();

  // Check the form values
  var reportName = form.reportName;
  var reportNamePublishing = form.reportNamePublishing;

  if (reportName === "" && reportNamePublishing === '') {
    ui.alert('You must enter a sheet report name and publishing report name.');
  } else if (reportName === '' && reportNamePublishing !== '') {
    ui.alert('You must enter a sheet report name.');
  } else if (reportName !== '' && reportNamePublishing === '') {
    ui.alert('You must enter a publishing report name.')
  } else {


    var reportName = form.reportName;
    var reportNamePublishing = form.reportNamePublishing;

    // Authenticate API call
    var token = getToken(clientID, clientSecret);

    // Get OCLC numbers
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var lastrow = sheet.getLastRow() - 1;
    var range = sheet.getRange(2, 1, lastrow, 1);
    var oclcNumbers = range.getValues();

    ss.toast('Creating report ... This might take a while. Wait for the next alert saying the report is ready.', 'Alert');

    // Set variable for array to receive new rows as they are created through the loop
    var dataNew = [];

    // Set variable for array to receive uniquely held number
    var totalUnique = [];

    // Set variable for array to receive number without holdings set
    var totalHoldingsNotSet = [];

    for (let rowNumber = 0; rowNumber < oclcNumbers.length; rowNumber++) {
      var oclcNumber = oclcNumbers[rowNumber];

      // Call the APIs
      var responseHoldings = callAPIHoldings(oclcNumber, token);
      var responseBibInfo = callAPIBibInfo(oclcNumber, token);
      // There was a problem, so stop
      if (responseHoldings === null || responseBibInfo === null)
        return;

      // Turn the JSON string into JSON object
      var recordHoldings = JSON.parse(responseHoldings).briefRecords[0];
      var recordBib = JSON.parse(responseBibInfo);

      // Return OCLC numbers that do not have AS0 holdings set
      if (recordHoldings.institutionHolding.briefHoldings != null && recordHoldings.institutionHolding.briefHoldings[0].oclcSymbol !== 'AS0') {
        totalHoldingsNotSet.push(oclcNumber);
      }

      // Remove oclcSymbol institutionType === 'OTHER' from array
      if (recordHoldings.institutionHolding.totalHoldingCount >= 1) {
        var adjustedHoldings = [];
        for (x = 0; x < recordHoldings.institutionHolding.briefHoldings.length; x++) {
          var holdingLibrary = recordHoldings.institutionHolding.briefHoldings[x];
          if (holdingLibrary.institutionType != 'OTHER') {
            adjustedHoldings.push(holdingLibrary);
          }
        }
      }

      // Calculate the number of unique titles
      if (adjustedHoldings.length === 1) {
        totalUnique.push(1);
      }

      var photobook = '';
      var subjects = recordBib.subjects;
      if (subjects != null) {
        for (x = 0; x < subjects.length; x++) {
          var subject = recordBib.subjects[x];
          if (subject.subjectName.text === 'Photobooks.') {
            var photobook = 'Photobook';
            break;
          }
        }
      }

      var journal = '';
      if (recordBib.format.generalFormat === 'Jrnl') {
        var journal = 'Journal';
      }

      var aucCat = '';
      if (recordBib.note.eventNotes != null && recordBib.note.eventNotes.length == 2) {
        aucCat = "Auction Catalog";
      }

      var format = photobook + journal + aucCat;
      var formatStyle = SpreadsheetApp.newTextStyle()
        .setBold(true)
        .setFontFamily('Roboto')
        .setForegroundColor('#f89497')
        .build();
      var formatStyled = SpreadsheetApp.newRichTextValue()
        .setText(format)
        .setTextStyle(formatStyle)
        .build();

      var cdlcURL = 'http://libweb.lib.tcu.edu/F?func=find-b&local_base=mus_acm&search_code=WRD&request=' + oclcNumber;
      var titleColonRemoved = recordHoldings.title.replace(/\s:/g, ':');
      var titleSemicolonRemoved = titleColonRemoved.replace(/\s;/g, ';');
      var titleFinalPeriodRemoved = titleSemicolonRemoved.replace(/\.$/, '');
      var title = titleCaps(titleFinalPeriodRemoved);
      var titleStyle = SpreadsheetApp.newTextStyle()
        .setItalic(true)
        .setBold(false)
        .setFontFamily('Roboto')
        .setFontSize(10)
        .setForegroundColor('#6f6f6f')
        .setUnderline(false)
        .build();
      var titleStyled = SpreadsheetApp.newRichTextValue()
        .setText(title)
        .setLinkUrl(cdlcURL)
        .setTextStyle(titleStyle)
        .build();

      var author = '';
      if (recordHoldings.creator != null) {
        var author = recordHoldings.creator;
      }
      var authorStyle = SpreadsheetApp.newTextStyle()
        .setBold(false)
        .setFontFamily('Roboto')
        .setForegroundColor('#747474')
        .build();
      var authorStyled = SpreadsheetApp.newRichTextValue()
        .setText(author)
        .setTextStyle(authorStyle)
        .build();

      var publisher = '';
      if (recordHoldings.publisher != null) {
        var publisherClean = recordHoldings.publisher.replace(/\s;/g, ';');
        var publisher = publisherClean + ', ' + recordHoldings.date;
      }
      var publisherStyle = SpreadsheetApp.newTextStyle()
        .setBold(false)
        .setFontFamily('Roboto')
        .setForegroundColor('#747474')
        .build();
      var publisherStyled = SpreadsheetApp.newRichTextValue()
        .setText(publisher)
        .setTextStyle(publisherStyle)
        .build();

      if (adjustedHoldings.length > 1) {
        var nearestHoldingLibraryName = adjustedHoldings[1].institutionName;
        var nearestHoldingLibraryCity = adjustedHoldings[1].address.city;
        var holdingsInfo = 'We are one of ' + adjustedHoldings.length + ' libraries worldwide that have this title. The next nearest library that has it is ' + nearestHoldingLibraryName + ', ' + nearestHoldingLibraryCity + '.';
      } else {
        // var holdingsInfo = 'The Carter Library is the only library in the world that has this title.';
        var holdingsInfo = 'Only held at the Carter Library!';
      }
      var holdingsInfoStyle = SpreadsheetApp.newTextStyle()
        .setBold(false)
        .setFontFamily('Roboto')
        .setForegroundColor('#b88c62')
        .build();
      var holdingsInfoStyled = SpreadsheetApp.newRichTextValue()
        .setText(holdingsInfo)
        .setTextStyle(holdingsInfoStyle)
        .build();

      /*       var imageURL = '=IMAGE('https://paynegap.info/media/book-images/2021-02/' + 'oclcNumber' + '.jpg');
            var image = 'Stay tuned for an image ...';
            var imageStyle = SpreadsheetApp.newTextStyle()
              .setBold(false)
              .build();
            var imageStyled = SpreadsheetApp.newRichTextValue()
              .setText(image)
              .setTextStyle(imageStyle)
              .build();
       */

      /*       var cdlcURL = 'http://libweb.lib.tcu.edu/F?func=find-b&local_base=mus_acm&search_code=PRVID&request=' + oclcNumber;
            var cdlcURL = 'http://libweb.lib.tcu.edu/F?func=find-b&local_base=mus_acm&search_code=WRD&request=' + oclcNumber;
            var cdlcLinkStyle = SpreadsheetApp.newTextStyle()
              .setUnderline(false)
              .setForegroundColor('#747474')
              .build();
            var cdlcLinkStyled = SpreadsheetApp.newRichTextValue()
              .setText('CDLC Details')
              .setLinkUrl(cdlcURL)
              .setTextStyle(cdlcLinkStyle)
              .build();
       */
      data = [titleStyled, authorStyled, publisherStyled, formatStyled, holdingsInfoStyled];

      // Add the new row to the array
      dataNew.push(data);
    }

    var reportSheet = ss.insertSheet(reportName);

    // Report intro
    var reportIntroStyle = SpreadsheetApp.newTextStyle()
      .setUnderline(false)
      .setFontFamily('Roboto')
      .setItalic(false)
      .setFontSize(16)
      .setBold(false)
      .setForegroundColor('#4a9a87')
      .build();
    var reportIntroURLStyle = SpreadsheetApp.newTextStyle()
      .setUnderline(false)
      .setFontFamily('Roboto')
      .setItalic(false)
      .setFontSize(16)
      .setBold(true)
      .setForegroundColor('#4a9a87')
      // .setLinkUrl('https://www.cartermuseum.org/research-carter/library')
      .build();
      /* var reportIntroACMAAlibStyled = SpreadsheetApp.newRichTextValue()
      .setText('The Amon Carter Museum Research Library')
      .setLinkUrl('https://www.cartermuseum.org/research-carter/library')
      .setTextStyle(reportIntroStyle)
      .build(); */
    var reportIntroMainTextStyled = SpreadsheetApp.newRichTextValue()
      .setText('The Amon Carter Museum of American Art Research Library added ' + oclcNumbers.length + ' titles to its collection in ' + reportNamePublishing + '. As a collection focusing on American art and photography, we often collect titles scarcely available in other libraries. We share these titles through interlibary loan and by making them available to any visitor to our Reading Room.')
      .setTextStyle(reportIntroStyle)
      .setTextStyle(4,55,reportIntroURLStyle)
      .build();

    // Calculate percentage of uniquely held titles and style the data
    var uniquePercentage = totalUnique.length / oclcNumbers.length * 100;
    var uniquePercentageFormatted = String(uniquePercentage).replace(/\..*/, '');
    var uniquePercentageStyle = SpreadsheetApp.newTextStyle()
      .setUnderline(false)
      .setFontFamily('Roboto')
      .setItalic(false)
      .setFontSize(100)
      .setBold(false)
      .setForegroundColor('#5f903e')
      .build();
    var uniquePercentageStyled = SpreadsheetApp.newRichTextValue()
      .setText(uniquePercentageFormatted + '%')
      .setTextStyle(uniquePercentageStyle)
      .build();

    //  of the titles in this report are uniquely held by the Carter Library
    var formattedDate = Utilities.formatDate(new Date(), "GMT", "MMMM, d, yyyy");
    var uniquePercentageStatementStyle = SpreadsheetApp.newTextStyle()
      .setUnderline(false)
      .setFontSize(17)
      .setFontFamily('Roboto')
      .setItalic(false)
      .setForegroundColor('#ffae5d')
      .build();
    var uniquePercentageStatementStyled = SpreadsheetApp.newRichTextValue()
      .setText('of the ' + oclcNumbers.length + ' titles in this report are uniquely held by the Carter Library according to WorldCat (as of reporting date, ' + formattedDate + ').')
      .setTextStyle(uniquePercentageStatementStyle)
      .build();

    //  report header footnote
    var reportFootNoteStyle = SpreadsheetApp.newTextStyle()
      .setUnderline(false)
      .setFontSize(10)
      .setFontFamily('Roboto')
      .setItalic(false)
      .setForegroundColor('#767676')
      .build();
    var reportFootNoteStyled = SpreadsheetApp.newRichTextValue()
      .setText('Each title in the report is a link to our catalog, which provides full information, including call number.\n\n817.989.5040\nlibrary@cartermuseum.org\nwww.cartermuseum.org/research-carter/library')
      .setTextStyle(reportFootNoteStyle)
      .build();

    reportSheet.insertRowBefore(1);
    reportSheet.getRange(1, 1, 1, 5).setWrap(true);
    reportSheet.getRange(1, 1).setVerticalAlignment('middle').setRichTextValue(reportIntroMainTextStyled);
    reportSheet.getRange(1, 2).setVerticalAlignment('middle').setRichTextValue(uniquePercentageStyled);
    reportSheet.getRange(1, 3, 1, 1).setVerticalAlignment('middle').setRichTextValue(uniquePercentageStatementStyled);
    reportSheet.getRange(1, 5).setVerticalAlignment('middle').setRichTextValue(reportFootNoteStyled);
    reportSheet.getRange(1, 1, 1, 5).setBackground('#f3f3f3');
    reportSheet.setRowHeight(1, 290);
    // reportSheet.setRowHeights(3, oclcNumbers.length, 30);
    reportSheet.setColumnWidth(1, 370);
    reportSheet.setColumnWidth(2, 250);
    reportSheet.setColumnWidth(3, 260);
    reportSheet.setColumnWidth(4, 125);
    reportSheet.setColumnWidth(5, 375);

    reportSheet.insertRowAfter(1);
    reportSheet.setRowHeight(2, 50);
    reportSheet.getRange(2, 1, 1, 5)
      .setValues([['Title (a-z)', 'Author', 'Publisher', 'Special Type', 'Holding Info']])
      .setBackground('#f3f3f3')
      .setVerticalAlignment('middle')
      .setFontFamily('Roboto')
      .setFontColor('#6f6f6f')
      .setFontSize(11)
      .setFontWeight('bold')
      .setBorder(false, false, true, false, false, false, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    reportSheet.getRange(3, 1, oclcNumbers.length, 5)
      .setWrap(true)
      .setBackground('#f3f3f3')
      .setVerticalAlignment('top')
      .setBorder(true, false, true, false, false, true, '#666666', SpreadsheetApp.BorderStyle.SOLID)
      .setRichTextValues(dataNew)
      .sort(1);
    // .applyRowBanding();

    reportSheet.getRange(1, 3, 1, 2).merge();

    reportSheet.setHiddenGridlines(true);

    if (totalHoldingsNotSet.length > 0) {
      var lastRow = reportSheet.getLastRow();
      reportSheet.insertRowAfter(lastRow + 1);
      reportSheet.getRange(lastRow + 2,1,1,totalHoldingsNotSet.length).setValues([totalHoldingsNotSet]);
    }

    ss.toast('Report is ready!', 'Alert');
  }
}
