Option Explicit

Function ModifiedFileCopy()

    Dim objFS
    Dim strCopyFrom
    Dim strCopyTo
    Dim lngCopyDays
    Dim blnCopyCreate
    Dim blnCopyUpdate
    Dim blnCopyAccess
    Dim dteTargetDate
    Dim strCopyFilename
    Dim strCopyToFilename
    Dim blnCopyLockOn
    Dim strFileName
    Dim strList
    Dim lngCopyCount
    Dim blnSubFolder
    Dim lngSwitchCount
    Dim dteDateCreated
    Dim dteDateLastModified
    Dim dteDateLastAccessed

    'プロシージャの結果を初期化します
    ModifiedFileCopy = False

    'コマンドライン引数でコピー元ファイル名とコピー先フォルダ名が
    '指定されているかをチェックします
    If WScript.Arguments.Unnamed.Count = 2 Or _
       WScript.Arguments.Unnamed.Count = 3 Then

       'コピー元フォルダの有無をチェックします
       If FolderCheck(WScript.Arguments.Unnamed(0), strCopyFrom) = False Then
          '見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : コピー元フォルダは存在しません。 " & WScript.Arguments.Unnamed(0)
          Exit Function
       End If

       'コピー先フォルダの有無をチェックします
       If FolderCheck(WScript.Arguments.Unnamed(1), strCopyTo) = False Then
          '見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : コピー先フォルダは存在しません。 " & WScript.Arguments.Unnamed(1)
          Exit Function
       End If

       lngCopyDays = 0

       '日数をチェック
       If WScript.Arguments.Unnamed.Count = 3 Then

          If IsNumeric(WScript.Arguments.Unnamed(2)) = True Then
             lngCopyDays = CLng(WScript.Arguments.Unnamed(2))

             If lngCopyDays < 0 Then
                '0以下を指定したときはエラーメッセージを表示しプロシージャを終了します
                WScript.Echo "ERROR : 日数は0以上で指定してください。 " & WScript.Arguments.Unnamed(2)
                Exit Function
             End If

          End If

       End If

    Else
       'パラメタが正しく指定されていないときは、エラーメッセージを表示し
       'プロシージャを終了します
       WScript.Echo "ERROR : パラメタを正しく指定してください。"
       Exit Function
    End If

    'スイッチ数を取得
    lngSwitchCount = WScript.Arguments.Named.Count

    '/c /u /aスイッチのチェック
    ':作成日で選別スイッチ
    If WScript.Arguments.Named.Exists("c") = False Then
       blnCopyCreate = False
    Else
       blnCopyCreate = True
       lngSwitchCount = lngSwitchCount - 1
    End If
    ':更新日で選別スイッチ
    If WScript.Arguments.Named.Exists("u") = False Then
       blnCopyUpdate = False
    Else
       blnCopyUpdate = True
       lngSwitchCount = lngSwitchCount - 1
    End If
    ':最終アクセス日で選別スイッチ
    If WScript.Arguments.Named.Exists("a") = False Then
       blnCopyAccess = False
    Else
       blnCopyAccess = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/subスイッチのチェック
    If WScript.Arguments.Named.Exists("sub") = False Then
       blnSubFolder = False
    Else
       blnSubFolder = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/c /u /aのいずれも指定されていないときは従来仕様を指定したことにする
    If blnCopyCreate = False And _
       blnCopyUpdate = False And _
       blnCopyAccess = False Then

       blnCopyCreate = True
       blnCopyUpdate = True

    End If

    'スイッチ種類のチェック
    If lngSwitchCount > 0 Then '余分なスイッチがある
       'プロシージャを終了します
       WScript.Echo "ERROR : 無効なスイッチがあります。"
       Exit Function
    End If

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'コピーファイルのカウント用
    lngCopyCount = 0

    'ファイルリストの取得
    If blnSubFolder = False Then
       strList = SearchFolder(strCopyFrom)     '/sub未指定時
    Else
       strList = sSearchFolderAll(strCopyFrom) '/sub指定時
    End If

    dteTargetDate = CStr(DateAdd("d", (lngCopyDays * (-1)), Date))

    '[POINT!]ファイルリストの処理
    For Each strFileName In strList

        '[POINT!]ファイルのタイムスタンプを取得する
        On Error Resume Next 'ZIPやLZHのときに実行時エラーを発生させないようにする
        dteDateCreated = vbNullString
        dteDateCreated = FormatDateTime(objFS.GetFile(strFileName).DateCreated, vbShortDate)
        dteDateLastModified = vbNullString
        dteDateLastModified = FormatDateTime(objFS.GetFile(strFileName).DateLastModified, vbShortDate)
        dteDateLastAccessed = vbNullString
        dteDateLastAccessed = FormatDateTime(objFS.GetFile(strFileName).DateLastAccessed, vbShortDate)
        On Error GoTo 0

        blnCopyLockOn = False

        '[POINT!]対象のファイルかどうかを判断します。
        If blnCopyCreate  = True          And _
           dteDateCreated = dteTargetDate Then
           blnCopyLockOn = True
        End If

        If blnCopyUpdate       = True          And _
           dteDateLastModified = dteTargetDate Then
           blnCopyLockOn = True
        End If

        If blnCopyAccess       = True          And _
           dteDateLastAccessed = dteTargetDate Then
           blnCopyLockOn = True
        End If

        'コピー対象ファイルのとき
        If blnCopyLockOn = True Then
           'フルパスからファイル名を取得します
           strCopyFilename = objFS.GetFilename(strFileName)
           'コピー先ファイル名を編集します
           strCopyToFilename = objFS.BuildPath(strCopyTo, strCopyFilename)
           'ファイルをコピーします
           objFS.CopyFile strFileName, strCopyToFilename
           WScript.Echo "コピーしました " & strFileName
           lngCopyCount = lngCopyCount + 1
        End If

    Next

    '処理結果メッセージの表示
    If lngCopyCount > 0 Then
       WScript.Echo lngCopyCount & "個のファイルをコピーしました。"
    Else
       WScript.Echo "コピーするファイルはありませんでした。"
    End If

    'オブジェクトの破棄
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    ModifiedFileCopy = True

End Function
