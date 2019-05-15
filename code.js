function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var values = sheet.getRange('A:A').getValues().filter(String);
  
  var count = sheet.getRange(1, 2).getValue();
  
  if (count > values.length) {
    Browser.msgBox('人数が足りません');
    return;
  }

  var result = lottery(values, count);
  sheet.getRange('E:E').clearContent();
  sheet.getRange(1, 5).setValue('月曜日');
  sheet.getRange(2, 5, result.length, 1).setValues(result);
  
  result = lottery(values, count);
  sheet.getRange('F:F').clearContent();
  sheet.getRange(1, 6).setValue('火曜日');
  sheet.getRange(2, 6, result.length, 1).setValues(result);
  
  result = lottery(values, count);
  sheet.getRange('G:G').clearContent();
  sheet.getRange(1, 7).setValue('水曜日');
  sheet.getRange(2, 7, result.length, 1).setValues(result);
  
  result = lottery(values, count);
  sheet.getRange('H:H').clearContent();
  sheet.getRange(1, 8).setValue('木曜日');
  sheet.getRange(2, 8, result.length, 1).setValues(result);
  
  result = lottery(values, count);
  sheet.getRange('I:I').clearContent();
  sheet.getRange(1, 9).setValue('金曜日');
  sheet.getRange(2, 9, result.length, 1).setValues(result);
}

function lottery(values, count) {
  var result = [];
  
  for (var i = 0; i < count; i++) {
    var value = values[Math.floor(Math.random() * values.length)];
    
    var array = Array.prototype.concat.apply([], result);
    if (array.indexOf(value) !== -1) {
      i--;
      continue;
    }
    
    result.push([value]);
  }
  
  return result;
}