Option Explicit

Const cWait = 500 '監視間隔間隔(100～1000を推奨)

Function TailCommand()

    Dim objFS
    Dim strTargetFile
    Dim lngExec
    Dim objInFile
    Dim strData
    Dim dteStartTime

    'プロシージャの結果を初期化します
    TailCommand = False

    'コマンドライン引数でターゲットファイルが
    '指定されているかをチェックします
    If WScript.Arguments.Count = 2 Then

       'ターゲットファイルの有無をチェックします
       If FileCheck(WScript.Arguments(0), strTargetFile) = False Then
          '見つからなかったときはエラーメッセージを表示し
          'プロシージャを終了します
          WScript.Echo "ERROR : 指定したファイルは存在しません。 " & WScript.Arguments(0)
          Exit Function
       End If

       '処理秒数パラメタを取得
       lngExec = WScript.Arguments(1)

       '秒数が数値かどうかをチェックします
       If IsNumeric(lngExec) = True Then

          lngExec = CLng(lngExec) '念のため整数型に変換します

          '秒数が0未満のときはパラメタエラーでプロシージャを終了
          If lngExec < 0 Then
             WScript.Echo "ERROR : 秒数には0以上を指定してください。"
             Exit Function
          End If
       Else
          WScript.Echo "ERROR : 秒数は数値で指定してください。"
          Exit Function
       End If

    Else
       'パラメタが正しく指定されていないときは、エラーメッセージを表示し
       'プロシージャを終了します
       WScript.Echo "ERROR : パラメタを正しく指定してください。"
       Exit Function
    End If

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    dteStartTime = Now()

    '対象ファイルの準備
    Set objInFile = objFS.OpenTextFile(strTargetFile, ForReading)
    'カラファイルでない場合は全読み込みを行う
    If objInFile.AtEndOfStream = False Then
       objInFile.ReadAll
    End If

    Do

       If objInFile.AtEndOfStream = False Then
          strData = objInFile.ReadAll
          WScript.Stdout.Write strData
       End If

       WScript.Sleep(cWait) '0.5秒待機

    Loop Until DateDiff("s", dteStartTime, Now()) >= lngExec And _
               lngExec > 0

    objInFile.Close

    'オブジェクトの破棄
    Set objInFile = Nothing
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    TailCommand = True

End Function
