/**
 * googleスプレッドシートの各種シートに、javascriptからアクセスするための<br>
 * ナビゲーション情報シートを生成します<br>
 * @author ms32
 * @version 1.0.0
 * 
 */

var spreadsheetWriteInfo = {
    // メニュー表示文字列（トップメニュー）
    MENU_STR_TOP:        "sheetsNavigationGenerator",
    // メニュー表示文字列（サブメニュー）
    MENU_STR_SUB:        "run",
    
    /**
     * 全シートのgidを収集して、スプレッドシートの先頭シートに書き込みます
     * @param {spreadsheetObj} 対象スプレッドシート
     */
    create: function(spreadsheetObj){
        "use strict";
        
        var WRITE_ROW_START_POS = 1;        // シート内の書き込み開始位置(row)
        var WRITE_COL_START_POS = 1;        // シート内の書き込み開始位置(column)
        
        var spreadsheetId = spreadsheetObj.getId(), 
            spreadsheetURL = spreadsheetObj.getUrl(), 
            spreadsheetName = spreadsheetObj.getName(), 
            sheets = spreadsheetObj.getSheets(), 
            i = 0, iLength = 0, 
            sheetObj = null, range = null, 
            writeArray = [];

        // 先頭には、このスプレッドシートにアクセスするためのkeyを書き込んでおきます
        writeArray.push(["key", spreadsheetId]);
        
        // 2列目以降には、各シートのシート名とポインタ情報を書き込みます
        for (i = 0, iLength = sheets.length; i < iLength; i+=1) {
            sheetObj = sheets[i];
            writeArray.push([sheetObj.getName(), sheetObj.getSheetId()]);  // 名前とシートID(gid)のペアを出力
        }

        // 集めた内容を先頭シートに出力
        sheetObj = sheets[0];
        sheetObj.clear();
        range = sheetObj.getRange(WRITE_ROW_START_POS, WRITE_COL_START_POS, writeArray.length, writeArray[0].length);
        range.setValues(writeArray);
    },
    
    /**
     * グローバルコールバックのonOpenから呼び出されます
     * メニュー表示と関数呼び先を設定します
     */
    onOpen: function(spreadsheetObj){
        "use strict";

        var menuList = [];
        menuList.push({
            name : spreadsheetWriteInfo.MENU_STR_SUB,
            functionName : "onMenuRun"
        });
        spreadsheetObj.addMenu(spreadsheetWriteInfo.MENU_STR_TOP, menuList); 
        
    }
    
};


/**
 * システムから呼ばれるグローバルコールバック 
 * シートオープン時に呼ばれる
 */
function onOpen() {
    "use strict";
   var spreadsheetObj = SpreadsheetApp.getActiveSpreadsheet();
   spreadsheetWriteInfo.onOpen(spreadsheetObj);
}

/**
 * メニュー実行から呼ばれるグローバルコールバック 
 * メニュー選択時に呼ばれる
 */
function onMenuRun() {
    "use strict";
    var spreadsheetObj = SpreadsheetApp.getActiveSpreadsheet();
    return spreadsheetWriteInfo.create(spreadsheetObj);
}
