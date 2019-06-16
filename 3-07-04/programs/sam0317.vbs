Option Explicit

Function CopySelectType()

    Dim objWSArg
    Dim objFS
    Dim strCopyFrom
    Dim strCopyFilename
    Dim strCopyFileExt
    Dim strFileName
    Dim strList
    Dim lngCopyCount
    Dim strExtName
    Dim blnSubFolder
    Dim lngSwitchCount
    Dim fn 

    'プロシージャの結果を初期化します
    CopySelectType = False

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'objWSArgオブジェクトのインスタンスをコピー
    Set objWSArg = WScript.Arguments

    'コマンドライン引数でコピー元ファイル名とコピー先フォルダ名が
    '指定されているかをチェックします
    If objWSArg.Unnamed.Count = 1 Then

       '検索元フォルダの有無をチェックします
       If FolderCheck(objWSArg.Unnamed(0), strCopyFrom) = False Then
          '見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : 検索元フォルダは存在しません。 " & objWSArg.Unnamed(0)
          Exit Function
       End If

    Else
       'パラメタが正しく指定されていないときは、エラーメッセージを表示し
       'プロシージャを終了します
       WScript.Echo "ERROR : パラメタを正しく指定してください。"
       Exit Function
    End If

    lngSwitchCount = WScript.Arguments.Named.Count

    '拡張子指定の引数
    If objWSArg.Named.Exists("e") = False Then
       WScript.Echo "ERROR : 拡張子を /e:??? で指定してください。 "
       Exit Function
    Else
       '抽出条件（拡張子）を取得
       strExtName = objWSArg.Named("e")

       '必須入力チェック
       If strExtName = vbNullString Then
          WScript.Echo "ERROR : 拡張子が指定されていません。"
          Exit Function
       End If

       lngSwitchCount = lngSwitchCount - 1

    End If

    '/subスイッチのチェック
    If objWSArg.Named.Exists("sub") = False Then
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

    'コピーファイルのカウント用
    lngCopyCount = 0

    'ファイルリストの取得
    If blnSubFolder = False Then
       strList = SearchFolder(strCopyFrom)     '/sub未指定時
    Else
       strList = sSearchFolderAll(strCopyFrom) '/sub指定時
    End If

    '[POINT!]ファイルリストの処理
    For Each strFileName In strList

        'ファイル名を取得します
        strCopyFilename = objFS.GetFilename(strFileName)

        '拡張子を取得します
        strCopyFileExt = objFS.GetExtensionName(strFileName)

        '[POINT!]拡張子が一致するかどうかを判断します
        '        このときに大文字で判断します
        If InStr(1,strCopyFilename,strExtName,vbTextCompare) > 0 Then
           'コピー元ファイル名を検索します
           Set fn = objFS.GetFile(strFileName) 
           WScript.Echo "名前：" & fn.Name & "------更新日時：" & fn.DateLastModified & "------サイズ：" & FormatNumber(fn.Size, 0)
           WScript.Echo "パス:" & strFileName

           lngCopyCount = lngCopyCount + 1
        End If

    Next

    '処理結果メッセージの表示
    If lngCopyCount > 0 Then
       WScript.Echo lngCopyCount & "個のファイルを検索しました。"
    Else
       WScript.Echo "検索するファイルはありませんでした。"
    End If

    'オブジェクトの破棄
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    CopySelectType = True

End Function
