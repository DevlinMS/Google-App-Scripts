function sheetsGUI() {
  /*
  *This creates a new Scipts dropdown in the menu bar
  *Clicking on the 'Get item count' entree will execute the getCount script
  */
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts')
    .addItem('Get item count', 'getCount').addToUi();
  return 0;
}



function getCount() {
  /*
  *This function will start at cell B1, and take note of how many times
  *each unique value is seen. The tracking ends five columns over, column F
  *It will go down to the last row that has a value
  *It ignores any empty cell
  *Data is stored in a dictionary as item: count
  */
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
  //Calls funtion to write values
  writeValues(dictToSortedArray(equipment));
  return 0;
}


function writeValues(info) {
  /*
  *This function writes the dictionary info onto collumns H and I
  *With the key on H and the value on I
  *Titles are written on H1 and I1
  *The data goes to the row equal to the lenght of the dict + 1
  */
  const sheet = SpreadsheetApp.getActiveSheet();
  //Write titles
  sheet.getRange(1,8).setValue('Equipment').setFontSize(14).setFontFamily('Calibri').setBackground('#D3D3D3');
  sheet.getRange(1,9).setValue('Amount').setFontSize(14).setFontFamily('Calibri').setBackground('#D3D3D3');
  //Write Data
  let count = 2;
  for (let i = 0; i < info.length; i++) {
    sheet.getRange(count,8).setValue(info[i][0]).setFontSize(14).setFontFamily('Calibri');
    sheet.getRange(count,9).setValue(info[i][1]).setFontSize(14).setFontFamily('Calibri');
    count += 1;
  }
  //clear 25 rows below the last written entree
  //This is incase the amount of unique entrees has decreased since
  //The last time the script was run
  sheet.getRange(count,8,25,2).setValue(null);
  //resize the new collumns
  sheet.autoResizeColumns(8, 2);
  return 0;
}

function dictToSortedArray(obj) {
  /*
  *Sorts a dictionary alphabetically 
  */
  const unsorted = Object.entries(obj);
  const sorted = unsorted.sort();
  Logger.log(sorted);
  return sorted;
}

