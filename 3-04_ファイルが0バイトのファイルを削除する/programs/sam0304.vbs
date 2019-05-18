Option Explicit

Function DeleteZerobyte()

    Dim objFS
    Dim strExecFolder
    Dim strFileName
    Dim strList
    Dim lngDelCount
    Dim lngFileSize
    Dim blnSubFolder
    Dim lngSwitchCount

    'プロシージャの結果を初期化します
    DeleteZerobyte = False

    'コマンドライン引数でコピー元ファイル名とコピー先フォルダ名が
    '指定されているかをチェックします
    If WScript.Arguments.Unnamed.Count = 1 Then

       '処理フォルダの有無をチェックします
       If FolderCheck(WScript.Arguments.Unnamed(0), strExecFolder) = False Then
          '見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : 指定したフォルダは存在しません。 " & WScript.Arguments.Unnamed(0)
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

    '削除ファイルのカウント用
    lngDelCount = 0

    'ファイルリストの取得
    'ファイルリストの取得
    If blnSubFolder = False Then
       strList = SearchFolder(strExecFolder)     '/sub未指定時
    Else
       strList = sSearchFolderAll(strExecFolder) '/sub指定時
    End If

    '[POINT!]ファイルリストの処理
    For Each strFileName In strList

        '[POINT!]ファイルサイズを求める
        On Error Resume Next 'ZIPやLZHのときに実行時エラーを発生させないようにする
        lngFileSize = (-1)
        lngFileSize = objFS.GetFile(strFileName).Size
        On Error GoTo 0

        '[POINT!]ファイルサイズが０であれば、そのファイルを削除します
        If lngFileSize = 0 Then
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
    DeleteZerobyte = True

End Function
