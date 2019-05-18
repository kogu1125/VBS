Option Explicit

Function LogOutput()

    Dim objWSArg
    Dim objFS
    Dim strExecFolder
    Dim strFilePath
    Dim strFilename
    Dim strFilenameAfter
    Dim strNewFilename
    Dim strNewFile
    Dim strList
    Dim strListItem
    Dim strBefore
    Dim strAfter
    Dim blnNull
    Dim lngCount
    Dim strLogFile
    Dim strMsg

    lngCount = 0

    'プロシージャの結果を初期化します
    LogOutput = False

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'ログファイルの場所
    strLogFile = objFS.BuildPath(getCurrentDir(), "sam0315.log")

    'objWSArgオブジェクトのインスタンスをコピー
    Set objWSArg = WScript.Arguments

    '置き換え後がnullの場合はスイッチで行う
    If objWSArg.Named.Count = 1 Then

       'スイッチが正しく指定されている場合
       If objWSArg.Named.Exists("null") = True Then

          'カットの指定時は置き換え後をvbNullStringに
          strAfter = vbNullString

       'スイッチを間違えて指定した場合
       Else
          strMsg = "ERROR : /null 以外のスイッチが指定されています。"
          Call TextWriteBottom(strLogFile, strMsg)
          Exit Function
       End If

       blnNull = True

    Else
       blnNull = False
    End If

    'コマンドライン引数でターゲットフォルダと置き換え内容が
    '指定されているかをチェックします
    If objWSArg.Unnamed.Count = 3 Or _
       (blnNull = True And objWSArg.Unnamed.Count = 2)   Then

       '処理フォルダの有無をチェックします
       If FolderCheck(WScript.Arguments(0), strExecFolder) = False Then
          '見つからなかったときはエラーメッセージを表示し
          'プロシージャを終了します
          strMsg = "ERROR : 指定したフォルダは存在しません。 " & objWSArg.Unnamed(0)
          Call TextWriteBottom(strLogFile, strMsg)
          Exit Function
       End If

       strBefore =objWSArg.Unnamed(1) '置き換え前

       '置き換え後は/nullスイッチが無い場合のみ有効
       If blnNull = False Then
          strAfter = objWSArg.Unnamed(2) '置き換え後
       End If

       '置き換え前後が同じ場合は処理できない
       If UCase(strBefore) = UCase(strAfter) Then
          strMsg = "ERROR : 置き換え前後の内容が同じです。 "
          Call TextWriteBottom(strLogFile, strMsg)
          Exit Function
       End If

    Else
       'パラメタが正しく指定されていないときは、エラーメッセージを表示し
       'プロシージャを終了します
       strMsg = "ERROR : パラメタを正しく指定してください。"
       Call TextWriteBottom(strLogFile, strMsg)
       Exit Function
    End If

    'ファイルリストの取得
    strList = SearchFolder(strExecFolder)

    '[POINT!]ファイルリストの処理
    For Each strListItem In strList

        '取り出したアイテムから、ファイル名のみを取得
        strFilename = objFS.GetFilename(strListItem)

        'リネームの対象ファイルかどうかを判断します
        '対象ファイルであればファイル名をリネームします
        If InStr(1, strFilename, strBefore, vbTextCompare) > 0 Then

           lngCount = lngCount + 1 'サンプル改造用（通番カウント）

           'リネーム後のファイル名とパスを編集します
           strNewFilename = Replace(strFilename, strBefore, strAfter, 1, -1, vbTextCompare)
           strNewFile = objFS.BuildPath(strExecFolder, strNewFilename)

           '同じ名前のファイルがあるときは処理を中段する
           If objFS.FileExists(strNewFile) = True Then
              strMsg = "同名のファイルが存在します -> " & strNewFilename
              Call TextWriteBottom(strLogFile, strMsg)
           Else
              'FileMoveメソッドを使ってファイル名を変更します
              objFS.MoveFile strListItem, strNewFile

              strMsg = "ファイル名を変更しました 前: " & strFilename & _
                       " 後: " & strNewFilename
              Call TextWriteBottom(strLogFile, strMsg)

           End If

        End If

    Next

    strMsg = "処理が完了しました"
    Call TextWriteBottom(strLogFile, strMsg)

    'オブジェクトの破棄
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    LogOutput = True

End Function
