
function main() {
  const sheet = SpreadsheetApp.getActiveSheet();
  // メンバーリスト
  const memberArray = sheet.getRange('A:A').getValues().filter(String).map(member => member[0]);
  // 1日あたりの当番人数
  const dutyCount = sheet.getRange(2, 2).getValue();
  // 当番回数の上限
  const dutyLimit = sheet.getRange(2, 3).getValue();

  // { 'メンバーの名前', 当番回数 }
  let memberObjects = {};
  for (let member of memberArray) {
    memberObjects[member] = 0;
  }

  // 当番人数に対して、メンバーが足りない場合エラー
  if (dutyCount > memberArray.length) {
    Browser.msgBox('人数が足りません');
    return;
  }

  const result = lotteries(memberArray, dutyCount);
  sheet.getRange('E:E').clearContent();
  sheet.getRange(1, 5).setValue('月曜日');
  sheet.getRange(2, 5, result.length, 1).setValues(convertToArray2d(result));

  /*
  var array = [];
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      array.push(data[i][j]);
    }
  }

  var members = {};
  for (var i in array) {
    members[array[i]] = 0;
  }

  var memberCount = sheet.getRange(1, 2).getValue();

  if (memberCount > Object.keys(members).length) {
    Browser.msgBox('人数が足りません');
    return;
  }

  var result = lottery(members, memberCount);
  sheet.getRange('E:E').clearContent();
  sheet.getRange(1, 5).setValue('月曜日');
  sheet.getRange(2, 5, result.length, 1).setValues(result);

  result = lottery(members, memberCount);
  sheet.getRange('F:F').clearContent();
  sheet.getRange(1, 6).setValue('火曜日');
  sheet.getRange(2, 6, result.length, 1).setValues(result);

  result = lottery(members, memberCount);
  sheet.getRange('G:G').clearContent();
  sheet.getRange(1, 7).setValue('水曜日');
  sheet.getRange(2, 7, result.length, 1).setValues(result);

  result = lottery(members, memberCount);
  sheet.getRange('H:H').clearContent();
  sheet.getRange(1, 8).setValue('木曜日');
  sheet.getRange(2, 8, result.length, 1).setValues(result);

  result = lottery(members, memberCount);
  sheet.getRange('I:I').clearContent();
  sheet.getRange(1, 9).setValue('金曜日');
  sheet.getRange(2, 9, result.length, 1).setValues(result);
  */
}

/*
function lottery(members, memberCount) {
  var result = [];
  var limit = Math.round(memberCount * 5 / Object.keys(members).length);

  while (result.length < memberCount) {
    // 規定数未満登場しているメンバーのリスト
    var values = [];
    for (var i in Object.keys(members)) {
      if (members[Object.keys(members)[i]] < limit) {
        values.push(Object.keys(members)[i]);
      }
    }

    // 抽選結果と重複しているメンバーを、抽選元から除外
    for (var i = 0; i < values.length; i++) {
      if (arrayExists(result, values[i])) {
        values.splice(i, 1);
      }
    }

    // 抽選
    var value = values[Math.floor(Math.random() * values.length)];

    Logger.log("members: " + members);
    Logger.log(value);

    members[value]++;
    result.push([value]);
  }

  return result;
}
*/

/**
 * 渡した配列からcount回数抽選して返す関数
 * 抽選内容の重複は無し
 *
 * @param {Array.<string>} values
 * @param {number} count
 * @returns {Array.<string>}
 */
function lotteries(values, count) {
  if (values.length < count) {
    throw new Error("抽選回数に対して、データが足りません");
  }

  const lotteryBox = values.slice(0, values.length);
  const result = [];

  for (let i = 0; i < count; i++) {
    const lot = lottery(lotteryBox);

    for (let j = 0; j < lotteryBox.length; j++) {
      if (lot == lotteryBox[j]) {
        lotteryBox.splice(j, 1);
      }
    }

    result.push(lot);
  }

  return result;
}

/**
 * 渡した配列から1件ランダムに抽選して返す関数
 *
 * @param {Array.<string>} values
 * @returns {string}
 */
function lottery(values) {
  return values[Math.floor(Math.random() * values.length)];
}

/**
 * 渡した配列の中に、渡した変数と同じ値のものが存在した場合はtrue
 * そうでない場合false
 *
 * @param {Array.<string>} values
 * @param {string} target
 * @returns {boolean}
 */
function arrayExists(values, target) {
  for (const value of values) {
    if (value == target) {
      return true;
    }
  }

  return false;
}

/**
 * 1次元配列を2次元配列にして返す関数
 *
 * @param {Array.<*>} array
 * @returns {Array.<*>[]}
 */
function convertToArray2d(array) {
  const array2d = [];
  for (const value of array) {
    array2d.push([value]);
  }
  return array2d;
}
