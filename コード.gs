//2023清掃賞計算表
function onEdit(e) {

  var sheetN = e.source.getSheetName();
  //シート名が生データじゃなかったらスルー
  if (sheetN != "生データとシフト") { return; }
  var sheet = e.source.getActiveSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var bgc = sheet.getRange(row, col).getBackground();
  //編集されたセルが灰色だった場合スルー
  if (bgc == "#b7b7b7") { return; }
  //値の変更が無ければスルー
  if (e.value == e.oldValue) { return; }
  //（１，１）でなければスルー
  if (row != 1 || col != 1) { return; }

  if (e.value) {//TRUEでなければスルー
    sheet.getRange(1, 1).setValue(false);//falseに戻す

    var sheetary = sheet.getRange(2, 6, sheet.getLastRow() - 1, 2).getValues();//１４日以上にも対応
    var result = [["スタッフ名", "獲得ポイント", "獲得金額（円）"]];//結果配列
    Logger.log(sheetary);

    //二次元配列の行と列を入れ替える関数
    const transpose = a => a[0].map((_, c) => a.map(r => r[c]));

    for (var r = 0; r <= sheetary.length - 1; r++) {

      var members = sheetary[r][1].split("、");
      var point = sheetary[r][0] / members.length;
      Logger.log(members + "\n" + point);

      for (var rr = 0; rr <= members.length - 1; rr++) {

        var result_t = transpose(result);//行列入れ替え（第１列→第１行へ）
        var index = result_t[0].indexOf(members[rr]);//完全一致で検索（行列入れ替えたresultの１行目）

        if (index != -1) {//氏名追加済みの場合
          result[index][1] = result[index][1] + point;

        } else {//氏名未追加の場合
          result.push([members[rr], point]);

        }
      }

    }

    //金額計算
    for (var r = 1; r <= result.length - 1; r++) {
      result[r][2] = Math.round(result[r][1] * 100);
    }
    //金額で並べ替え
    result.sort((a, b) => { return b[2] - a[2]; });
    Logger.log(result);

    //結果シートをクリアして書き込み
    var sheet_result = e.source.getSheetByName("結果");
    sheet_result.clear();
    sheet_result.getRange(1, 1, result.length, result[0].length).setValues(result);

  }

}
