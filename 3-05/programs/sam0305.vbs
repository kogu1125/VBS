Option Explicit

Dim cDelimiter

'[POINT!]区切り文字
cDelimiter = vbTab

Function FolderSearch()

    Dim objFS
    Dim strExecFolder
    Dim varFilename
    Dim strPrintData
    Dim strList
    Dim blnSubFolder
    Dim lngSwitchCount
    Dim blnHeader
    Dim blnFileName
    Dim blnFileSize
    Dim blnFileCreate
    Dim blnFileUpdate
    Dim blnFileAccess
    Dim blnZipFile
    Dim varFileSize
    Dim dteDateCreated
    Dim dteDateLastModified
    Dim dteDateLastAccessed

    'プロシージャの結果を初期化します
    FolderSearch = False

    'コマンドライン引数に検索対象フォルダが指定されているとき
    If WScript.Arguments.Unnamed.Count = 1 Then

       '処理フォルダの有無をチェックします
       If FolderCheck(WScript.Arguments.Unnamed(0), strExecFolder) = False Then
          '見つからなかったときはエラーメッセージを表示し
          'プロシージャを終了します
          WScript.Echo "ERROR : 指定したフォルダは存在しません。 " & WScript.Arguments.Unnamed(0)
          Exit Function
       End If

    Else
       '検索対象フォルダを省略したときは、カレントディレクトリを検索
       strExecFolder = getCurrentDir()
    End If

    'スイッチ数を取得
    lngSwitchCount = WScript.Arguments.Named.Count

    '/t スイッチ（ヘッダー行）のチェック
    If WScript.Arguments.Named.Exists("t") = False Then
       blnHeader = False
    Else
       blnHeader = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/f スイッチ（ファイル名のみ表示）のチェック
    If WScript.Arguments.Named.Exists("f") = False Then
       blnFileName = False
    Else
       blnFileName = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/s スイッチ（ファイルサイズ表示）のチェック
    If WScript.Arguments.Named.Exists("s") = False Then
       blnFileSize = False
    Else
       blnFileSize = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/c /u /a(作成・更新・アクセス日)スイッチのチェック
    ':作成日で選別スイッチ
    If WScript.Arguments.Named.Exists("c") = False Then
       blnFileCreate = False
    Else
       blnFileCreate = True
       lngSwitchCount = lngSwitchCount - 1
    End If
    ':更新日で選別スイッチ
    If WScript.Arguments.Named.Exists("u") = False Then
       blnFileUpdate = False
    Else
       blnFileUpdate = True
       lngSwitchCount = lngSwitchCount - 1
    End If
    ':最終アクセス日で選別スイッチ
    If WScript.Arguments.Named.Exists("a") = False Then
       blnFileAccess = False
    Else
       blnFileAccess = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/zip (書庫ファイル内の検索)スイッチのチェック
    If WScript.Arguments.Named.Exists("zip") = False Then
       blnZipFile = False
    Else
       blnZipFile = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/sub (サブフォルダ検索)スイッチのチェック
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

    'ファイル一覧を取得します
    If blnSubFolder = False Then
       strList = SearchFolder(strExecFolder)     '/sub未指定時
    Else
       strList = sSearchFolderAll(strExecFolder) '/sub指定時
    End If

    'ヘッダー行の出力
    If blnHeader = True Then

       strPrintData = "ファイル名"

       'ファイルサイズ
       If blnFileSize = True Then
          strPrintData = strPrintData _
                       & cDelimiter _
                       & "サイズ"
       End If
       'ファイル作成日
       If blnFileCreate = True Then
          strPrintData = strPrintData _
                       & cDelimiter _
                       & "作成日"
       End If
       'ファイル更新日
       If blnFileUpdate = True Then
          strPrintData = strPrintData _
                       & cDelimiter _
                       & "更新日"
       End If
       'ファイル最終アクセス日
       If blnFileAccess = True Then
          strPrintData = strPrintData _
                       & cDelimiter _
                       & "最終アクセス日"
       End If

       WScript.Echo strPrintData

    End If

    '[POINT!]ファイルリストの処理
    For Each varFilename In strList

        strPrintData = vbNullString

        'パスも表示するかファイル名のみを表示するか
        If blnFileName = True Then
           strPrintData = strPrintData _
                        & objFS.GetFilename(varFilename)
        Else
           strPrintData = strPrintData _
                            & varFilename
        End If

        'ファイルサイズを取得
        '[POINT!]取得できないとき（varFileSizeがvbNullStringのとき)は書庫ファイル内のファイル
        On Error Resume Next
        varFileSize = vbNullString
        varFileSize = objFS.GetFile(varFilename).Size
        On Error GoTo 0

        If blnFileSize = True Then
           strPrintData = strPrintData _
                        & cDelimiter _
                        & varFileSize
        End If

        'ファイル作成日
        If blnFileCreate = True Then
           On Error Resume Next
           dteDateCreated = vbNullString
           dteDateCreated = objFS.GetFile(varFilename).DateCreated
           On Error GoTo 0
           strPrintData = strPrintData _
                        & cDelimiter _
                        & dteDateCreated
        End If

        'ファイル更新日
        If blnFileUpdate = True Then
           On Error Resume Next
           dteDateLastModified = vbNullString
           dteDateLastModified = objFS.GetFile(varFilename).DateLastModified
           On Error GoTo 0
           strPrintData = strPrintData _
                        & cDelimiter _
                        & dteDateLastModified
        End If

        'ファイル最終アクセス日
        If blnFileAccess = True Then
           On Error Resume Next
           dteDateLastAccessed = vbNullString
           dteDateLastAccessed = objFS.GetFile(varFilename).DateLastAccessed
           On Error GoTo 0
           strPrintData = strPrintData _
                        & cDelimiter _
                        & dteDateLastAccessed
        End If

        'ファイル情報を表示
        ':ファイルサイズが取得できた分(zip内のファイル以外)のみ表示
        ':ただし、/zipスイッチがあるときは、無条件に表示
        If varFileSize <> vbNullString Or _
          (blnZipFile  = True         ) Then

           WScript.Echo strPrintData

        End If

    Next

    'オブジェクトの破棄
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    FolderSearch = True

End Function
