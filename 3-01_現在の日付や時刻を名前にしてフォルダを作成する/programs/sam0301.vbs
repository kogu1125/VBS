Option Explicit

Function CreateFolder()

    Dim objWSArg
    Dim objFS
    Dim strFolderName
    Dim strCreatePath

    'プロシージャの結果を初期化します
    CreateFolder = False

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'objWSArgオブジェクトのインスタンスをコピー
    Set objWSArg = WScript.Arguments

    'コマンドライン引数で作成先が指定されているかをチェックします
    If objWSArg.Unnamed.Count > 0 Then

       'パスのチェック。見つからないときはプロシージャを終了
       If FolderCheck(objWSArg.Unnamed(0), strCreatePath) = False Then
          WScript.Echo "ERROR : パスが存在しません。 " & objWSArg.Unnamed(0)
          Exit Function
       End If

    Else
       'コマンドライン引数が指定されていないときはカレントディレクトリが標準パス
       strCreatePath = getCurrentDir()
    End If

    '[POINT!]パラメタの内容によって日付や時刻を出力します
    If objWSArg.Named.Count = 0 Then ' スイッチを省略したときはyyyymmdd
       strFolderName = Right("000" & Year(Date), 4) & _
                       Right("00" & Month(Date), 2) & _
                       Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D1") = True Then ' yyyymmdd
           strFolderName = Right("000" & Year(Date), 4) & _
                           Right("00" & Month(Date), 2) & _
                           Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D2") = True Then ' yyyymm
           strFolderName = Right("000" & Year(Date), 4) & _
                           Right("00" & Month(Date), 2)

    ElseIf objWSArg.Named.Exists("D3") = True Then ' mmdd
           strFolderName = Right("00" & Month(Date), 2) & _
                           Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D4") = True Then ' yyyy
           strFolderName = Right("000" & Year(Date), 4)

    ElseIf objWSArg.Named.Exists("D5") = True Then ' mm
           strFolderName = Right("00" & Month(Date), 2)

    ElseIf objWSArg.Named.Exists("D6") = True Then ' dd
           strFolderName = Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("T1") = True Then 'hhmmss
           strFolderName = Right("00" & Hour(Time), 2) & _
                           Right("00" & Minute(Time), 2) & _
                           Right("00" & Second(Time), 2)

    ElseIf objWSArg.Named.Exists("T2") = True Then ' hhmm
           strFolderName = Right("00" & Hour(Time), 2) & _
                           Right("00" & Minute(Time), 2)

    ElseIf objWSArg.Named.Exists("T3") = True Then ' mmss
           strFolderName = Right("00" & Minute(Time), 2) & _
                           Right("00" & Second(Time), 2)

    ElseIf objWSArg.Named.Exists("T4") = True Then ' yyyymm
           strFolderName = Right("00" & Hour(Time), 2)

    ElseIf objWSArg.Named.Exists("T5") = True Then ' mm
           strFolderName = Right("00" & Minute(Time), 2)

    ElseIf objWSArg.Named.Exists("T6") = True Then ' ss
           strFolderName = Right("00" & Second(Time), 2)

    Else
       '不明な書式スイッチはエラー
       WScript.Echo "ERROR : 書式を正しく指定してください。"
       Exit Function
    End If

    strCreatePath = objFS.BuildPath(strCreatePath, strFolderName)

    '[POINT!]フォルダを作成します
    '        ただし、既に同じ名前のフォルダがあるときは作成しません
    If objFS.FolderExists(strCreatePath) = False Then
       objFS.CreateFolder strCreatePath
       WScript.Echo "フォルダを作成しました"
       WScript.Echo strCreatePath
    Else
       WScript.Echo "ERROR : 同じ名前のフォルダが存在します"
       WScript.Echo strCreatePath
    End If

    'オブジェクトの破棄
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    CreateFolder = True

End Function
