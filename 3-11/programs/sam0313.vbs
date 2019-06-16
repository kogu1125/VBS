Option Explicit

Function FlatToFlat()

    Dim objFS
    Dim objCon
    Dim strInFile
    Dim strInFilename
    Dim strInPath
    Dim strOutFile
    Dim strOutFilename
    Dim objOutFile
    Dim objRS
    Dim strSQL

    'プロシージャの結果を初期化します
    FlatToFlat = False

    'コマンドライン引数で入力・出力ファイルが指定されているかを
    'チェックします
    If WScript.Arguments.Count >= 2  Then

       '入力の有無をチェックします
       If FileCheck(WScript.Arguments(0), strInFile) = False Then
          '見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : 入力ファイルが存在しません。 " & WScript.Arguments(0)
          Exit Function
       End If

       '出力の有無をチェックします
       If FileCheck(WScript.Arguments(1), strOutFile) = False Then
          '見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : 出力ファイルが存在しません。 " & WScript.Arguments(1)
          Exit Function
       End If

    Else
       'パラメタが正しく指定されていないときは、エラーメッセージを表示し
       'プロシージャを終了します
       WScript.Echo "ERROR : パラメタを正しく指定してください。"
       Exit Function
    End If

    'ADOコネクションを作成します
    strInPath = getFilePath(strInFile)
    If getTextConnection(objCon, strInPath) = False Then
       'コネクションが作成できなかったときはプロシージャを終了します
       WScript.Echo "ERROR : コネクションが作成できませんでした。"
       Exit Function
    End if

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '出力ファイルをクリアします
    Set objOutFile = objFS.OpenTextFile(strOutFile, ForWriting)
    objOutFile.Close

    '入力・出力ファイル名の取得
    strInFilename = objFS.GetFilename(strInFile)
    strOutFilename = objFS.GetFilename(strOutFile)

    '入力ファイルの読み込み
    Set objRS = objCon.Execute("select * from " & strInFilename)

    Do Until objRS.Eof = True

       '書き込み用SQLの編集
       strSQL = vbNullString
       strSQL = strSQL & "INSERT  INTO FLATOUTDATA.txt"
       strSQL = strSQL & "       (日付"
       strSQL = strSQL & "       ,担当者"
       strSQL = strSQL & "       ,商品コード"
       strSQL = strSQL & "       ,価格 ) "
       strSQL = strSQL & "VALUES ('" & objRS("日付") & "'"
       strSQL = strSQL & "       ,'" & objRS("担当者") & "'"
       strSQL = strSQL & "       ,'" & objRS("商品コード") & "'"
       strSQL = strSQL & "       ,'" & Right("0000" & objRS("価格"), 5) & "' )"

       '出力ファイルへ書き込み
       objCon.Execute(strSQL)

       '次のレコードを読む
       objRS.MoveNext

    Loop

    'コネクションとレコードセットのクローズ
    objRS.Close
    objCON.Close

    WScript.echo "処理が完了しました"

    'オブジェクトの破棄
    Set objRS = Nothing
    Set objCon = Nothing
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    FlatToFlat = True

End Function
