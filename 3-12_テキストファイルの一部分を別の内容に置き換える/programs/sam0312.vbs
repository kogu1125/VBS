Option Explicit

Function TextReplace()

    Dim objFS
    Dim strTargetFile
    Dim strBefore()
    Dim strAfter()
    Dim i
    Dim j
    Dim objInFile
    Dim strData
    Dim strCtrlA
    Dim strCtrlB
    Dim objNewFile
    Dim strNewFilePath
    Dim strNewFile

    'プロシージャの結果を初期化します
    TextReplace = False

    'コマンドライン引数の数3以上、かつ奇数かどうかをチェックします。
    If WScript.Arguments.Count >= 3 And _
       IsEven(WScript.Arguments.Count) = False Then

       'ターゲットファイルの有無をチェックします
       If FileCheck(WScript.Arguments(0), strTargetFile) = False Then
          'パスが見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : ターゲットファイルが存在しません。 " & WScript.Arguments(0)
          Exit Function
       End If

       j = 0

       '置き換え情報の保存
       For i=2 To WScript.Arguments.Count Step 2
           ReDim Preserve strBefore(j)
           ReDim Preserve strAfter(j)
           strBefore(j) = WScript.Arguments(i -1)
           strAfter(j) = WScript.Arguments(i)
           j = j + 1
       Next

    Else
       'パラメタが正しく指定されていないときは、エラーメッセージを表示し
       'プロシージャを終了します
       WScript.Echo "ERROR : パラメタを正しく指定してください。"
       Exit Function
    End If

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '元ファイルを読み込みます
    Set objInFile = objFS.OpenTextFile(strTargetFile, ForReading)
    strData = objInFile.ReadAll
    objInFile.Close

    '置き換え処理を行います(For終了値は配列の大きさ)
    For i=0 To UBound(strBefore)

        '置き換え前文字列の編集
        strCtrlB = strBefore(i)
        If UCase(strCtrlB) = "/S" Then       '制御文字(半角空白)
           strCtrlB = " "
        ElseIf UCase(strCtrlB) = "/W" Then   '制御文字(全角空白)
           strCtrlB = "　"
        End If

        '置き換え後文字列の編集
        strCtrlA = strAfter(i)
        If UCase(strCtrlA) = "/S" Then       '制御文字(半角空白)
           strCtrlA = " "
        ElseIf UCase(strCtrlA) = "/W" Then   '制御文字(全角空白)
           strCtrlA = "　"
        ElseIf UCase(strCtrlA) = "/NUL" Then '制御文字(カラ文字)
           strCtrlA = vbNullString
        End If

        '置き換え処理
        strData = Replace(strData, strCtrlB, strCtrlA)

    Next

    '中間ファイルの作成
    strNewFilePath = getFilePath(strTargetFile)
    strNewFile = objFS.BuildPath(strNewFilePath, objFS.GetTempName)
    Set objNewFile = objFS.CreateTextFile(strNewFile, ForWriting)

    '中間ファイルへの書き込み
    objNewFile.Write strData
    objNewFile.Close

    '中間ファイルをターゲットファイルにする
    objFS.DeleteFile strTargetFile
    objFS.MoveFile strNewFile, strTargetFile

    WScript.echo "置き換えが完了しました " & strTargetFile

    'オブジェクトの破棄
    Set objInFile = Nothing
    Set objNewFile = Nothing
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    TextReplace = True

End Function
