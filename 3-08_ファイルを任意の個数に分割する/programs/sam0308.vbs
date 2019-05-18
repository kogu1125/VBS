Option Explicit

Function FileCutter()

    Dim objWSArg
    Dim objFS
    Dim objADOST_R
    Dim objADOST_W
    Dim strTargetFile
    Dim strCreatePath
    Dim strCreateName
    Dim strCreateFilename
    Dim lngDivCount
    Dim lngDivSerial
    Dim bytData

    Dim lngDivParm
    Dim lngDivSize

    'プロシージャの結果を初期化します
    FileCutter = False

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'objWSArgオブジェクトのインスタンスをコピー
    Set objWSArg = WScript.Arguments

    'コマンドライン引数で対象のファイルが指定されているかを
    'チェックします
    If objWSArg.Unnamed.Count = 1 Then

       '対象ファイルの有無をチェックします。
       If FileCheck(objWSArg.Unnamed(0), strTargetFile) = False Then
          '見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : 対象ファイルが存在しません。 " & objWSArg.Unnamed(0)
          Exit Function
       End If

    Else
       'パラメタが正しく指定されていないときは、エラーメッセージを表示し
       'プロシージャを終了します
       WScript.Echo "ERROR : パラメタを正しく指定してください。"
       Exit Function
    End If

    '分割数のチェック
    If objWSArg.Named.Count = 0 Then
       '分割数省略時は2個
       lngDivParm = 2

    'スイッチがあるのに分割数が無い場合
    ElseIf objWSArg.Named.Exists("c") = False Then
       WScript.Echo "ERROR : 分割数は /c:?? で指定してください。"
       Exit Function
    Else

       '分割数省略時は2個
       lngDivParm = objWSArg.Named("c")

       '分割数が数値以外の場合はエラー
       If IsNumeric(lngDivParm) = False Then
          WScript.Echo "ERROR : 分割数は数値で指定してください。"
          Exit Function
       ElseIf CLng(lngDivParm) < 2 Then
          WScript.Echo "ERROR : 分割数は2以上で指定してください。"
          Exit Function
       End If

       '念のため数値型に変換
       lngDivParm = CLng(lngDivParm)

    End If

    '分割後の大きさを求める(端数は切り上げ)
    lngDivSize = Fix(objFS.GetFile(strTargetFile).Size / lngDivParm + 0.99999999)

    '分割数よりファイルのバイト数が小さいときは分割できない
    If objFS.GetFile(strTargetFile).Size < lngDivParm Then
       Set objFS = Nothing
       WScript.Echo "ERROR : " & lngDivParm & "個には分割できません。"
       Exit Function
    End if

    'ADODB.Streamオブジェクトを作成します
    Set objADOST_W = Createobject("ADODB.Stream") 'READ用
    Set objADOST_R = Createobject("ADODB.Stream") 'WRITE用

    '分割数のカウント用
    lngDivSerial = 0
    lngDivCount  = 0

    strCreatePath = getFilePath(strTargetFile)

    '入力ファイルをADODB.Streamでオープンします
    objADOST_R.Type = adTypeBinary
    objADOST_R.Open
    objADOST_R.LoadFromFile strTargetFile
    objADOST_R.Position = 0

    'ファイルの分割処理
    Do Until objADOST_R.EOS = True

       '分割数のカウント
       lngDivSerial = lngDivSerial + 1
       lngDivCount  = lngDivCount + 1

       '[POINT!]分割後サイズ分を読み込みます
       bytData = objADOST_R.Read(lngDivSize)

       '読み込んだデータを分割後ファイルへ出力します
       objADOST_W.Open
       objADOST_W.Type = adTypeBinary
       objADOST_W.Write bytData

       '[POINT!]ファイル名には分割番号を付加します
       strCreateName = objFS.GetFilename(strTargetFile) & _
                       "." & Right("000" & lngDivSerial, 3) & ".div"
       strCreateFilename = objFS.BuildPath(strCreatePath, strCreateName)

       WScript.echo "分割 -> " & strCreateFilename

       '念のため出力ファイルを削除します
       On Error Resume Next
       objFS.DeleteFile strCreateFilename
       On Error GoTo 0

       'ファイルを出力します
       objADOST_W.SaveToFile strCreateFilename

       objADOST_W.Close

    Loop

    '入力ファイルをクローズ
    objADOST_R.Close

    '分割結果の表示
    If lngDivCount > 1 Then
       WScript.Echo lngDivCount & "個のファイルに分割しました。"
    End If

    'オブジェクトの破棄
    Set objADOST_R = Nothing
    Set objADOST_W = Nothing
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    FileCutter = True

End Function
