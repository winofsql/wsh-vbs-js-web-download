if Wscript.Arguments.Count = 0 then
    Wscript.Echo "httpget url [savepath]"
    Wscript.Quit
end if

' *****************************
' ダウンロード用のオブジェクト
' *****************************
Set objSrvHTTP = Wscript.CreateObject("Msxml2.ServerXMLHTTP")

' *****************************
' 第1引数は URL
' *****************************
strUrl = Wscript.Arguments(0)
if Wscript.Arguments.Count = 1 then
    ' 第2引数が無い場合は、URL の最後の部分
    ' ( カレントにダウンロード )
    aData = Split(strUrl,"/")
    strFile = aData(Ubound(aData))
else
    ' 第2引数がある場合はそれをローカルファイルとする
    strFile = Wscript.Arguments(1)
end if

' *****************************
' ダウンロード要求
' *****************************
on error resume next
Call objSrvHTTP.Open("GET", strUrl, False )
if Err.Number <> 0 then
    Wscript.Echo Err.Description
    Wscript.Quit
end if
objSrvHTTP.Send
if Err.Number <> 0 then
    ' おそらくサーバーの指定が間違っている
    Wscript.Echo Err.Description
    Wscript.Quit
end if
on error goto 0

if objSrvHTTP.status = 404 then
    Wscript.Echo "URL が正しくありません(404)"
    Wscript.Quit
end if

' *****************************
' バイナリデータ保存用オブジェクト
' *****************************
Set Stream = Wscript.CreateObject("ADODB.Stream")
Stream.Open
Stream.Type = 1	' バイナリ
' 戻されたバイナリをファイルとしてストリームに書き込み
Stream.Write objSrvHTTP.responseBody
' ファイルとして保存
Stream.SaveToFile strFile, 2
Stream.Close

