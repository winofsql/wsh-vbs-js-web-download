// ****************************
// 初期処理
// ****************************
if ( WScript.Arguments.length == 0 ) {
    WScript.Echo( "httpget url [savepath]" );
    WScript.Quit();
}

// ****************************
// ダウンロード用のオブジェクト
// ****************************
var http = new ActiveXObject("Msxml2.ServerXMLHTTP")

// ****************************
// 第1引数は URL
// ****************************
var file;
var url = WScript.Arguments(0);
if ( WScript.Arguments.length == 1 ) {
    // 第2引数が無い場合は、URL の最後の部分
    // ( カレントにダウンロード )
    var aData = url.split("/");
    file = aData[aData.length-1];
}
else {
    // 第2引数がある場合はそれをローカルファイルとする
    file = WScript.Arguments(1);
}

// ****************************
// ダウンロード要求
// ****************************
WScript.Echo( url );
http.open("GET", url, false );
http.send();
try {
} catch (error) {
    WScript.Echo( error.description );
    WScript.Quit();
}

if ( http.status == 404  ) {
    WScript.Echo( "URL が正しくありません(404)" );
    WScript.Quit();
}

// ****************************
// バイナリデータ保存用オブジェクト
// ****************************
var stream = new ActiveXObject("ADODB.Stream");
stream.Open();
stream.Type = 1	// バイナリ
// 戻されたバイナリをファイルとしてストリームに書き込み
stream.Write( http.responseBody );
// ファイルとして保存
stream.SaveToFile( file, 2 );
stream.Close

// ****************************
// ファイルの最後
// ****************************

