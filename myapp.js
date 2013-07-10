var MYAPP = {
    // スプレッドシートへのアクセスキー
    SPREADSHEET_KEY: "表示したいスプレッドシートのkey",
    // アクセスしたいシート名
    TARGET_SHEET_NAME: "表示したいシート",
    // スプレッドシートへアクセスするベースURL
    BASE_URL: "http://spreadsheets.google.com/tq?key=",
    // jqGrid表示DOM名
    JQGRID_TARGET_DOM_NAME: "#jqGridList",          // for sample
    
    // navigationシートから取得した各シートのURL情報
    navigationDict: {},
    
    /**--------------------------------------------
    * Navigationシート内容取得コールバック適用
    *---------------------------------------------*/
    sendNavigationSheetCallback: function(){
        "use strict";
        
        // Navigationシートは先頭固定のため「gid=0」にしています
        var url   = MYAPP.BASE_URL + MYAPP.SPREADSHEET_KEY + "&gid=0&pub=0",
            query = new google.visualization.Query(url);
            
        query.send(MYAPP.runNavigationSheetCallback); 
    },
    
    /**--------------------------------------------
    * Navigationシート内容取得コールバック実行
    * @param  {response}  google.visualization.Query 実行後の返り値
    * Navigationシートから各シートへのURL情報を取得
    *---------------------------------------------*/
    runNavigationSheetCallback: function(response){
        "use strict";
        
        var data = response.getDataTable(),
            url, row, rowLength, key, keyStr, value, query;
            
        if (!data || response.isError()) {
            alert(response.getMessage() + ':' + response.getDetailedMessage());
            return;
        }        
        // 行単位にデータを読み込んでシート情報を取得
        for (row = 0, rowLength = data.getNumberOfRows(); row < rowLength; row++){
            keyStr = String(data.getFormattedValue(row, 0));
            value  = data.getFormattedValue(row, 1);
            MYAPP.navigationDict[keyStr] = value;
        }
        
        // 指定したgoogleスプレッドシートを読み込む
        url = MYAPP._getSpreadsheetUrl(MYAPP.TARGET_SHEET_NAME);
        if(url){
            query  = new google.visualization.Query(url);
            query.send(MYAPP.runTargetSheetCallback);
        }
    },
    
    /**--------------------------------------------
    * 取得対象シート内容取得コールバック実行
    * @param  {response}  google.visualization.Query 実行後の返り値
    * データを読み込みたいシートから内容取得
    *---------------------------------------------*/
    runTargetSheetCallback: function(response){
        "use strict";
        
        var DISP_ROW_MAX = 100,
            data = response.getDataTable(),
            value, row, col, rowLength, colLength, colName, rowTempDict,
            colArray = [], rowArray = [];
            
        if (!data || response.isError()) {
            alert(response.getMessage() + ':' + response.getDetailedMessage());
            return;
        }
        // loop max cache
        rowLength = data.getNumberOfRows();
        colLength = data.getNumberOfColumns();
        
        // ヘッダ読み込み(列データ)
        for ( col = 0; col < colLength; col += 1) {
            value = data.getColumnLabel(col);
            colArray.push({ "name" : value, "label" : value,
                "sortable" : true, "width" : 100
            });
        }
               
        // 行データ読み込み,格納
        for ( row = 0; row < rowLength; row += 1) {
            rowTempDict = [];
            for ( col = 0; col < colLength; col += 1) {
                value   = data.getValue(row, col);          // 指定位置のデータ取得
                colName = data.getColumnLabel(col);         // 格納位置割り出し
                rowTempDict[colName] = value;               // 行データの1要素格納
            }
            rowArray.push(rowTempDict);                     // 行データ格納
        }
        
        // jqGrid表示設定
        jQuery(MYAPP.JQGRID_TARGET_DOM_NAME).jqGrid({
            data: rowArray,
            colModel: colArray,
            datatype: "local",
            height: 'auto',
            sortorder: "ASC",
            multiselect: false,
            rowNum:DISP_ROW_MAX,
            excel: true,
            viewrecords: true,
            caption: "[" + MYAPP.TARGET_SHEET_NAME + "] の内容"
        });
    },

    /**- - - - - - - - - - - - - - - - - - - - - - -
    * 指定したシート名のURLを返す
    * @param  {sheetName} スプレッドシートのシート名
    * @return {String}    URL文字列(見つからない場合はnull)
    *- - - - - - - - - - - - - - - - - - - - - - -*/
    _getSpreadsheetUrl: function(sheetName){
        
        "use strict";
        var url, pageId;
        
        // 対象のシート名からシートID(gid)を取得してURLを生成する
        if( sheetName in MYAPP.navigationDict ){
            pageId = String(MYAPP.navigationDict[sheetName]);
            url    = MYAPP.BASE_URL + MYAPP.SPREADSHEET_KEY + 
                    "&gid=" + pageId + "&pub=0";
            return url;
        }
        alert("ERROR:指定したスプレッドシートの名前が見つかりません (NAME:" + sheetName + ")");
        return null;
    }
    
};



