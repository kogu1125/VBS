Option Explicit

Function FileJoin()

    Dim objWSArg
    Dim objFS
    Dim objADOST_R
    Dim objADOST_W
    Dim strCopyFrom
    Dim strCopyTo
    Dim strCopyFilename
    Dim strCopyFileBasename
    Dim strCopyFileExt
    Dim strFileName
    Dim strList
    Dim lngCopyCount
    Dim strTempFilename
    Dim bytData
    Dim strJoinExt
    Dim blnSubFolder
    Dim lngSwitchCount

    'プロシージャの結果を初期化します
    FileJoin = False

    'objWSArgオブジェクトのインスタンスをコピー
    Set objWSArg = WScript.Arguments

    'コマンドライン引数でコピー元フォルダ名とコピー先ファイル名が
    '指定されているかをチェックします
    If objWSArg.Unnamed.Count = 2 Then

       'コピー元フォルダの有無をチェックします
       If FolderCheck(objWSArg.Unnamed(0), strCopyFrom) = False Then
          '見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : コピー元フォルダは存在しません。 " & objWSArg.Unnamed(0)
          Exit Function
       End If

       'コピー先フォルダの有無をチェックします
       'ただし、依頼パス情報にはファイル名も含みます
       If FileFolderCheck(objWSArg.Unnamed(1), strCopyTo) = False Then
          '見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : 作成先のパスが存在しません。 " & objWSArg.Unnamed(1)
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
       WScript.Echo "ERROR : 結合するファイルの拡張子を /e:?? で指定してください。"
       Exit Function
    Else
       '結合ファイルの拡張子を取得
       strJoinExt = objWSArg.Named("e")

       '必須入力チェック
       If strJoinExt = vbNullString Then
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

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")
    'ADODB.Streamオブジェクトを作成します
    Set objADOST_W = Createobject("ADODB.Stream") 'READ用
    Set objADOST_R = Createobject("ADODB.Stream") 'WRITE用

    'コピーファイルのカウント用
    lngCopyCount = 0

    'ファイルリストの取得
    If blnSubFolder = False Then
       strList = SearchFolder(strCopyFrom)     '/sub未指定時
    Else
       strList = sSearchFolderAll(strCopyFrom) '/sub指定時
    End If

    '中間ファイルを作成します
    strTempFilename = objFS.BuildPath(strCopyFrom, objFS.GetTempName)

    '出力ファイルをADODB.Stream(バイナリモード)で開きます
    objADOST_W.Open
    objADOST_W.Type = adTypeBinary

    '[POINT!]ファイルリストの処理
    For Each strFileName In strList

        'フルパスからファイル名を取得します
        strCopyFilename = objFS.GetFilename(strFileName)
        'フルパスからファイル名を取得します
        strCopyFileBasename = objFS.GetBasename(strFileName)
        '同じく、拡張子を取得します
        strCopyFileExt  = objFS.GetExtensionName(strFileName)

        '[POINT!]拡張子が一致するかどうかを判断します（大文字で判断）
        If UCase(strCopyFileExt) = UCase(strJoinExt) Then

           'ファイルADODB.Streamで読み込みます
           objADOST_R.Type = adTypeBinary
           objADOST_R.Open
           objADOST_R.LoadFromFile strFileName
           objADOST_R.Position = 0
           bytData = objADOST_R.Read()
           objADOST_R.Close

           'ファイルが読み込めたときは中間ファイルへ出力する準備をします
           If IsNull(bytData)=False Then
              objADOST_W.Write bytData 'バッファへ出力
           End If

           WScript.Echo "結合します : " & strFileName
           lngCopyCount = lngCopyCount + 1
        End If

    Next

    '結合ファイルの作成と、処理結果メッセージの表示
    If lngCopyCount > 0 Then

       WScript.Echo lngCopyCount & "個のファイルを結合しました。"

       '出力ファイルを保存します
       objADOST_W.SaveToFile strTempFilename
       objADOST_W.Close

       '中間ファイルを作成先にコピーし、中間ファイルを削除します
       objFS.CopyFile strTempFilename, strCopyTo
       objFS.DeleteFile strTempFilename

    Else
       WScript.Echo "結合するファイルはありませんでした。"
    End If

    'オブジェクトの破棄
    Set objADOST_R = Nothing
    Set objADOST_W = Nothing
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    FileJoin = True

End Function
