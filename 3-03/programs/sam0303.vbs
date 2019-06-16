Option Explicit

Function DeleteTimelimit()

    Dim objFS
    Dim strExecFolder
    Dim lngDays
    Dim lngFileDays
    Dim strFileName
    Dim strList
    Dim lngDelCount
    Dim blnSubFolder
    Dim lngSwitchCount

    'プロシージャの結果を初期化します
    DeleteTimelimit = False

    'コマンドライン引数でコピー元ファイル名とコピー先フォルダ名が
    '指定されているかをチェックします
    If WScript.Arguments.Unnamed.Count = 2 Then

       '処理フォルダの有無をチェックします
       If FolderCheck(WScript.Arguments.Unnamed(0), strExecFolder) = False Then
          '見つからなかったときはエラーメッセージを表示し
          'プロシージャを終了します
          WScript.Echo "ERROR : 指定したフォルダは存在しません。 " & WScript.Arguments.Unnamed(0)
          Exit Function
       End If

       '日数パラメタを取得
       lngDays = WScript.Arguments.Unnamed(1)

       '日数が数値かどうかをチェックします
       If IsNumeric(lngDays) = True Then

          lngDays = CLng(lngDays) '念のため整数型に変換します

          '日数が0未満のときはパラメタエラーでプロシージャを終了
          If lngDays < 0 Then
             WScript.Echo "ERROR : 日数には0以上を指定してください。"
             Exit Function
          End If
       Else
          WScript.Echo "ERROR : 日数は数値で指定してください。"
          Exit Function
       End If

    Else
       'パラメタが正しく指定されていないときは、エラーメッセージを表示し
       'プロシージャを終了します
       WScript.Echo "ERROR : パラメタを正しく指定してください。"
       Exit Function
    End If

    'スイッチ数を取得
    lngSwitchCount = WScript.Arguments.Named.Count

    '/subスイッチのチェック
    If WScript.Arguments.Named.Exists("sub") = False Then
       blnSubFolder = False
    Else
       blnSubFolder = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    'スイッチ種類のチェック
    If lngSwitchCount > 0 Then '余分なスイッチがある
       'プロシージャを終了します
       WScript.Echo "ERROR : 無効なスイッチがあります。"
       Exit Function
    End If

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    lngDelCount = 0 '削除ファイルのカウント用

    'ファイルリストの取得
    If blnSubFolder = False Then
       strList = SearchFolder(strExecFolder)     '/sub未指定時
    Else
       strList = sSearchFolderAll(strExecFolder) '/sub指定時
    End If

    '[POINT!]ファイルリストの処理
    For Each strFileName In strList

        '[POINT!]ファイル作成日からの経過日数を求める
        On Error Resume Next 'ZIPやLZHのときに実行時エラーを発生させないようにする
        lngFileDays = -1
        lngFileDays = DateDiff("d", objFS.GetFile(strFileName).DateCreated, Date)
        On Error GoTo 0

        '[POINT!]指定した期間を経過していれば、そのファイルを削除します
        If lngFileDays >= lngDays Then
           objFS.DeleteFile strFileName
           WScript.Echo "削除しました " & strFileName
           lngDelCount = lngDelCount + 1
        End If

    Next

    '処理結果メッセージの表示
    If lngDelCount > 0 Then
       WScript.Echo lngDelCount & "個のファイルを削除しました"
    Else
       WScript.Echo "削除できるファイルはありませんでした。"
    End If

    'オブジェクトの破棄
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    DeleteTimelimit = True

End Function
