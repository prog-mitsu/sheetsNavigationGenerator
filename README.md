sheetsNavigationGenerator
=========================

various sheet of google spreadsheet, it will generate a navigation information sheet to access from javascript.

# webからgoogleスプレッドシートの各種シートに自在にアクセスする

## はじめに
webサイト側から、googleスプレッドシートの複数のシートにアクセスしたいニーズがあったとします。
GASなら、以下のフローで済むと思います。

+ スプレッドシートオブジェクトにアクセス
+ 支配下のシート一覧を取得
+ それぞれ対象のシートにアクセス

が、javascriptでアクセスする場合は…という事で、 Google Visualization API を試みました。
URL決め打ちで「シート」にはアクセスできますが、「スプレッドシート」にアクセスする方法が見当たりませんでした。
(スプレッドシート支配下のシート一覧が取れない。  取れるといいんだけど…)

スプレッドシート側が静的でシート名もシートURLも固定、という場合は困りませんが、

+ webサイト側でシート一覧を表示したい
+ スプレッドシート側のシート数が増減したりして、シートのURLがしょっちゅう変わる
+ シートのURL決め打ちではなく、シート名からシートにアクセスしたい

という場合などに困ってしまいます。

結局、

1. **GASでスプレッドシート側の先頭シートに、各種シートへのナビゲーション情報を出力する**<br>(このシートのURLは固定する)
2. **JS側からはそのナビゲーションを参照 → 各種ページへのURLを取得してアクセスする…**

という美しくない解法で実現しました。

**ナビゲーション情報シートはこんなんです**
![2013-07-10_07h01_38.png](https://qiita-image-store.s3.amazonaws.com/0/9439/1b8537a7-6041-e018-ba96-1925f6d45c24.png)

色々メンドウだったので、主に未来の自分に向けてメモを残しておきます。

**※もっと簡単で良い方法があるはずです、だれか教えて下さい**
javascript単体でスプレッドシート(シートではなく)情報に直にアクセスする手軽な方法あるのかな？

全コード [github](https://github.com/prog-mitsu/sheetsNavigationGenerator)

以下は概要メモですので、シートへのアクセス詳細などは  [参考サイト](#section1) をご参照下さい。

## 内容について

### google apps script(googleスプレッドシート側)

各シート情報を集めて、先頭のシートに出力する処理を行います。
実行しやすいように、シートのメニューに「sheetsNavigationGenerator」→「run」 を追加しています。

** googleスプレッドシートの各種シートに、javascriptからアクセスするための**
** ナビゲーション情報シートを生成します**

```js:sheetsNavigationGenerator.gs

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

```

#### 実行イメージ ####
![2013-07-10_15h59_17.png](https://qiita-image-store.s3.amazonaws.com/0/9439/3b9f760f-633a-997d-97c4-a810d64360ef.png)


### web(JavaScript)側

Google Visualization APIでシート情報を取得しますが、

+ 先頭シート(便宜上Navigationシートと呼びます)にアクセス
+ Navigationシートから、目的のシート名のgidを取得してurlを生成します
+ 生成したurlから対象シートにアクセスし、データを取得します

こんなフローでデータ取得しています。
以下、javascriptでシートの内容を取得して、jqGridで内容を表示するサンプルです。

```js:myapp.js

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

```

##動作サンプル
[サンプルページ](https://sites.google.com/site/qiitatestpublic20130710/home/)を設置しました。

## 最後に
全コードはこちらです [github](https://github.com/prog-mitsu/sheetsNavigationGenerator)
スプレッドシート側のシートが増減しまくったりするケースで便利だと思います。
また、JavaScript側に決め打ちのURLをベタベタ書かなくて良いので、そこそこデータドリブンです。
(そもそも、JavaScript側からスプレッドシート情報取れればこんな回りくどいことしなくていいんだけど…)
大事なので２度書きますが、**※もっと簡単で良い方法があるはずです、だれか教えて下さい**

## 参考サイト<a name="section1">  
[Google Spreadsheets を簡易 SQL DB に！「Google Visualization API」 - WebOS Goodies ](http://webos-goodies.jp/archives/51310352.html)

