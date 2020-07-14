// @NotOnlyCurrentDoc

function submission(e) {
  let sheet = SpreadsheetApp.getActiveSheet();
  let lastColumn = sheet.getLastColumn();
  let values = e.namedValues;
  let keyVal = [];
  
  //find key and value for form submission and set to Array
  for (let i = 2; i <= lastColumn; i++) {
    let header = sheet.getRange(1, i).getValue();
    let value = values[header] !== undefined ? values[header] : '';
    keyVal.push([header, value]);
  }
  
  //Create PDF of Form Submission
    let destination = DriveApp.getFolderById(''); //Include drive Folder ID for the PDF to be stored in

    let html = '<body style="text-align: center; color: #20365F; margin: 0; padding: 0; font-size: 12px;">' +
                '<hr style="margin: 25px 0 15px 0; border: 1px solid #20365F;"  />' +
                '<div style="position: relative; margin: 0; padding: 0;">' +
                '<div style="font-size: 26px; margin: 0 0 35px 0;">TITLE</div>' + //Change title and subtitle
                '<div style="font-size: 26px; margin: 10px 0 35px 0;">SUBTITLE</div>' +
                '<img src=""' +  //If you wish to have a logo in the top right corner include it here as BASE64
                'style="width: 80px; position: absolute; right: 30px; top: 0;">' +
                '</div>' +
                '<div style="width:90%; margin: auto; text-align: left;">' +
                '<hr style="margin: 15px 0 10px 0; border: 1px solid #20365F;"  />' +
                '<table style="border: 1px solid #20365F; font-size: 13px;">' +
                  keyVal.map((x) => '<tr><td style="border: 1px solid #20365F; width: 50%; color: #248ec2; padding: 3px;"><strong>' + x[0] + '</strong></td><td style="border: 1px solid #20365F; padding: 3px;">' + x[1] + '</td></tr>') +
                '</table>' +
                '<hr style="margin: 35px 0 35px 0; border: 1px solid #20365F;"  />' +
                '<div style="margin-top: 15px; text-align: center;">Digitally signed: ' + new Date().toString() + '</div>' +
                '</div>' + 
                '</div>' +
                '</div>' +
                '<hr style="margin: 25px 0 10px 0; border: 1px solid #20365F;"  />' +
                '<div style="margin-top: 15px; text-align: center;"><strong>Private and Confidential</strong></div>' +
                '</body>';
    
    //Capitalised Title regardless of submission
    let name = values['Full name'][0].toString().trim().split(' '); //Choose form field with which to name the file - preset is 'Full name'
    let nameArr = [];
    name.map(x => {
      let first = x.charAt(0).toUpperCase();
      let remaining = x.slice(1);
      nameArr.push(first + remaining);
    })
    
    //create blob and pdf - then add to drive
    let blob = Utilities.newBlob(html, "text/html", nameArr.join(' ') + '.html');
    let pdf = blob.getAs("application/pdf");

    destination.createFile(pdf).setName(nameArr.join(' ') + ".pdf");
}
