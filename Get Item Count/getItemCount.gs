function sheetsGUI() {
  const ui = SpreadsheetApp.getUi();
  //create ui element
  ui.createMenu('Scripts')
    //add box to dropdown that activates get count function
    .addItem('Get item count', 'getCount').addToUi();
    return 0;
}



function getCount() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getRange(1,2,sheet.getLastRow(),5).getValues();
  let equipment = {};

  //loop through all entires in the sheet
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      
      
      //if an entree is empty, ignore it
      if (data[i][j] != '') {
        //if already in dict, increase count by one
        let item = data[i][j]
        if (equipment[item] >= 1) {
          equipment[item] = equipment[item] + 1;
        }
        //else add it to dict with count of one
        else {
          equipment[item] = 1;
        }
      }
    }
  }
  writeValues(dictToSortedArray(equipment));
  return 0;
}


function writeValues(info) {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1,8).setValue('Equipment').setFontSize(14).setFontFamily('Calibri').setBackground('#D3D3D3');
  sheet.getRange(1,9).setValue('Amount').setFontSize(14).setFontFamily('Calibri').setBackground('#D3D3D3');
  //start writing on second row
  let count = 2;
  for (let i = 0; i < info.length; i++) {
    sheet.getRange(count,8).setValue(info[i][0]).setFontSize(14).setFontFamily('Calibri');
    sheet.getRange(count,9).setValue(info[i][1]).setFontSize(14).setFontFamily('Calibri');
    count += 1;
  }
  //clear below written incase of fewer equipment types
  sheet.getRange(count,8,25,2).setValue(null);
  //resize collumns
  sheet.autoResizeColumns(8, 2);
  return 0;
}

function dictToSortedArray(obj) {
  const unsorted = Object.entries(obj);
  const sorted = unsorted.sort();
  Logger.log(sorted);
  return sorted;
}

