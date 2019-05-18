Option Explicit

Function TimestampCopy()

    Dim objWSArg
    Dim objFS
    Dim strCopyFrom
    Dim strCopyTo
    Dim strFileName
    Dim strPrefix

    'プロシージャの結果を初期化します
    TimestampCopy = False

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'objWSArgオブジェクトのインスタンスをコピー
    Set objWSArg = WScript.Arguments

    'コマンドライン引数でコピー元ファイル名とコピー先フォルダ名が
    '指定されているかをチェックします
    If objWSArg.Unnamed.Count = 2 Then

       'コピー元ファイルの有無をチェックします
       If FileCheck(objWSArg.Unnamed(0), strCopyFrom) = False Then
          WScript.Echo "コピーするファイル " & objWSArg.Unnamed(0) & " は、存在しません。"
          Exit Function
       End If

       'コピー先フォルダの有無をチェックします
       If FolderCheck(objWSArg.Unnamed(1), strCopyTo) = False Then
          'それでも見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "コピー先フォルダ " & objWSArg.Unnamed(1) & " は、存在しません。"
          Exit Function
       End If

    Else
       'パラメタが正しく指定されていないときは、エラーメッセージを表示し
       'プロシージャを終了します
       WScript.Echo "ERROR : パラメタを正しく指定してください。"
       Exit Function
    End If

    '[POINT!]追加する日付時刻の編集
    If objWSArg.Named.Count = 0 Then ' スイッチを省略したときはyyyymmdd
       strPrefix = Right("000" & Year(Date), 4) & _
                   Right("00" & Month(Date), 2) & _
                   Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D1") = True Then ' yyyymmdd
           strPrefix = Right("000" & Year(Date), 4) & _
                       Right("00" & Month(Date), 2) & _
                       Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D2") = True Then ' yyyymm
           strPrefix = Right("000" & Year(Date), 4) & _
                       Right("00" & Month(Date), 2)

    ElseIf objWSArg.Named.Exists("D3") = True Then ' mmdd
           strPrefix = Right("00" & Month(Date), 2) & _
                       Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D4") = True Then ' yyyy
           strPrefix = Right("000" & Year(Date), 4)

    ElseIf objWSArg.Named.Exists("D5") = True Then ' mm
           strPrefix = Right("00" & Month(Date), 2)

    ElseIf objWSArg.Named.Exists("D6") = True Then ' dd
           strPrefix = Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("T1") = True Then 'hhmmss
           strPrefix = Right("00" & Hour(Time), 2) & _
                       Right("00" & Minute(Time), 2) & _
                       Right("00" & Second(Time), 2)

    ElseIf objWSArg.Named.Exists("T2") = True Then ' hhmm
           strPrefix = Right("00" & Hour(Time), 2) & _
                       Right("00" & Minute(Time), 2)

    ElseIf objWSArg.Named.Exists("T3") = True Then ' mmss
           strPrefix = Right("00" & Minute(Time), 2) & _
                       Right("00" & Second(Time), 2)

    ElseIf objWSArg.Named.Exists("T4") = True Then ' yyyymm
           strPrefix = Right("00" & Hour(Time), 2)

    ElseIf objWSArg.Named.Exists("T5") = True Then ' mm
           strPrefix = Right("00" & Minute(Time), 2)

    ElseIf objWSArg.Named.Exists("T6") = True Then ' ss
           strPrefix = Right("00" & Second(Time), 2)

    Else
       '不明な書式スイッチはエラー
       WScript.Echo "ERROR : 書式を正しく指定してください。"
       Exit Function
    End If

    '[POINT!]コピー先のファイル名とパスを編集します(yyyymmdd_ファイル名)
    strFileName = strPrefix & "_" & objFS.GetFilename(strCopyFrom)
    strCopyTo = objFS.BuildPath(strCopyTo, strFileName)

    '[POINT!]ファイルをコピーします
    'ただし、既に同じ名前のファイルがあるときは作成しません
    If objFS.FileExists(strCopyTo) = False Then
       objFS.CopyFile strCopyFrom, strCopyTo
       WScript.Echo "ファイルをコピーしました"
       WScript.Echo strCopyTo
    Else
       WScript.Echo "同じ名前のファイルが存在します"
       WScript.Echo strCopyTo
    End If

    'オブジェクトの破棄
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    TimestampCopy = True

End Function
