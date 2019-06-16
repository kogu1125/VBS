Option Explicit

Function FileRename()

    Dim objWSArg
    Dim objFS
    Dim strExecFolder
    Dim strFilePath
    Dim strFilename
    Dim strFileBasename
    Dim strFileExt
    Dim strFilenameAfter
    Dim strNewFoldername
    Dim strNewFilename
    Dim strNewFile
    Dim strList
    Dim strListItem
    Dim strBefore
    Dim strAfter
    Dim blnNull
    Dim lngCount
    Dim blnSubFolder
    Dim lngSwitchCount

    'プロシージャの結果を初期化します
    FileRename = False

    'objWSArgオブジェクトのインスタンスをコピー
    Set objWSArg = WScript.Arguments

    lngSwitchCount = objWSArg.Named.Count

    '置き換え後をnullにする場合のスイッチ
    If objWSArg.Named.Exists("null") = False Then
       blnNull = False
    Else
       'カットの指定時は置き換え後をvbNullStringに
       strAfter = vbNullString
       blnNull  = True

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

    'コマンドライン引数でターゲットフォルダと置き換え内容が
    '指定されているかをチェックします
    If objWSArg.Unnamed.Count = 3 Or _
       (blnNull = True And objWSArg.Unnamed.Count = 2)   Then

       '処理フォルダの有無をチェックします
       If FolderCheck(WScript.Arguments(0), strExecFolder) = False Then
          '見つからなかったときはエラーメッセージを表示し
          'プロシージャを終了します
          WScript.Echo "ERROR : 指定したフォルダは存在しません。 " & objWSArg.Unnamed(0)
          Exit Function
       End If

       strBefore =objWSArg.Unnamed(1) '置き換え前

       '置き換え後は/nullスイッチが無い場合のみ有効
       If blnNull = False Then
          strAfter = objWSArg.Unnamed(2) '置き換え後
       End If

       '置き換え前後が同じ場合は処理できない
       If UCase(strBefore) = UCase(strAfter) Then
          WScript.Echo "ERROR : 置き換え前後の内容が同じです。 "
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

    'ファイルリストの取得
    If blnSubFolder = False Then
       strList = SearchFolder(strExecFolder)     '/sub未指定時
    Else
       strList = sSearchFolderAll(strExecFolder) '/sub指定時
    End If

    '変数の初期化
    lngCount = 0

    '[POINT!]ファイルリストの処理
    For Each strListItem In strList

        '取り出したアイテムから、フォルダのパスを取得
        strNewFoldername = getFilePath(strListItem)
        '取り出したアイテムから、ファイル名のみを取得
        strFilename      = objFS.GetFilename(strListItem)
        '拡張子を含まないファイル名（サンプル改造用）
        strFileBasename  = objFS.GetBasename(strListItem)
        '拡張子（サンプル改造用）
        strFileExt       = objFS.GetExtensionName(strListItem)

        'リネームの対象ファイルかどうかを判断します
        '対象ファイルであればファイル名をリネームします
        If InStr(1, strFilename, strBefore, vbTextCompare) > 0 Then

           lngCount = lngCount + 1 'サンプル改造用（通番カウント）

           'リネーム後のファイル名とパスを編集します
           strNewFilename = Replace(strFilename, strBefore, strAfter, 1, -1, vbTextCompare)
           strNewFile = objFS.BuildPath(strNewFoldername, strNewFilename)

           '同じ名前のファイルがあるときは処理を中段する
           If objFS.FileExists(strNewFile) = True Then
              WScript.echo "同名のファイルが存在します -> " & strNewFilename
           Else
              'FileMoveメソッドを使ってファイル名を変更します
              objFS.MoveFile strListItem, strNewFile

              WScript.echo "ファイル名を変更しました 前: " & strFilename & _
                           " 後: " & strNewFilename

           End If

        End If

    Next

    WScript.echo "処理が完了しました"

    'オブジェクトの破棄
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    FileRename = True

End Function
