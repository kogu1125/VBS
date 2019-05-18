Option Explicit

'==========================================================
'= フォルダの有無チェック
'==========================================================
Function FolderCheck(ByVal tmpPath, ByRef strPath)

    Dim objFS
    Dim strCheckPath

    'プロシージャの結果を初期化します
    FolderCheck = False

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'チェックするフォルダのパス情報を作成（絶対パスに編集）
    strCheckPath = objFS.GetAbsolutePathName(tmpPath)

    '編集したパスでフォルダの有無をチェック
    If objFS.FolderExists(strCheckPath) = True Then
       FolderCheck = True
       strPath = strCheckPath 'パス情報を返す
    End If

    Set objFS = Nothing 'オブジェクトの破棄

End Function

'==========================================================
'= ファイル名を含むパスの場合のフォルダの有無チェック
'==========================================================
Function FileFolderCheck(ByVal tmpPath, ByRef strPath)

    Dim objFS
    Dim objPath
    Dim strCheckPath
    Dim strTmpPath

    'プロシージャの結果を初期化します
    FileFolderCheck = False

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '絶対パスに編集
    strTmpPath = objFS.GetAbsolutePathName(tmpPath)

    'さらにファイルまでのフルパスからパス情報だけを取得
    strCheckPath = Replace(strTmpPath, objFS.GetFilename(strTmpPath), vbNullString)

    '編集したパスでフォルダの有無をチェック
    If objFS.FolderExists(strCheckPath) = True Then
       FileFolderCheck = True
       strPath = strTmpPath 'パス情報を返す
    End If

    Set objFS = Nothing 'オブジェクトの破棄

End Function

'==========================================================
'= ファイルリストの取得(指定フォルダのみ)
'==========================================================
Function SearchFolder(tmpExecFolder)

    Dim objApl
    Dim objFolder
    Dim objFolderItems
    Dim objItem
    Dim lngCount
    dim strFileList()

    lngCount = 0

    '[POINT!]Shellオブジェクトを作成します
    Set objApl = CreateObject("Shell.Application")

    '[POINT!]検索するフォルダのオブジェクトを作成します
    Set objFolder = objApl.Namespace(tmpExecFolder)

    '[POINT!]フォルダ内を検索します
    For Each objItem In objFolder.Items

        '[POINT!]取り出した物がファイルかどうかを判断します
        '[POINT!]zip書庫ファイルは「フォルダ」で認識される場合があります
        If objItem.IsFolder = False Then
           Redim Preserve strFileList(lngCount)
           strFileList(lngCount) = objItem.Path
           lngCount = lngCount + 1
        End If

    Next

    '戻り値にはファイルの一覧返します
    SearchFolder = strFileList

    'オブジェクトの破棄
    Set objItem = Nothing
    Set objFolderItems = Nothing
    Set objFolder = Nothing
    Set objApl = Nothing

End Function

'==============================================================
'= ファイルリストの取得(下の階層まで検索する)
'==============================================================
Function sSearchFolderAll(tmpExecFolder)

    Dim objApl
    Dim objFolder
    Dim strFileList()

    'Shellオブジェクトを作成します
    Set objApl = CreateObject("Shell.Application")

    '検索するフォルダのオブジェクトを作成します
    Set objFolder = objApl.Namespace(tmpExecFolder)

    'フォルダ検索処理を呼び出します
    Call sSearchFolderAll_Sub(objFolder.Items, strFileList)

    '戻り値にはファイルの一覧返します
    sSearchFolderAll = strFileList

    'オブジェクトの破棄
    Set objFolder = Nothing
    Set objApl = Nothing

End Function

'==============================================================
'= フォルダ内に含まれるファイルやフォルダを検索する(再帰呼び出し)
'= :sSearchFolderAllのサブルーチン
'==============================================================
Sub sSearchFolderAll_Sub(ByVal tmpFolderItems, ByRef tmpFileList)

    Dim objFolderItems
    Dim objItem
    Dim lngCount

    '配列の大きさを再度求める
    lngCount = 0
    On Error Resume Next
    lngCount = UBound(tmpFileList) + 1
    On Error Goto 0

    'フォルダ内を検索
    For Each objItem In tmpFolderItems

        '取り出した物がファイルかフォルダかを判定
        If objItem.IsFolder Then
           'フォルダであれば、Itemsオブジェクトを作り、
           'それを引数としてsSearchFolderAll_Subを「再帰呼び出し」します
           Set objFolderItems = objItem.GetFolder.Items
           Call sSearchFolderAll_Sub(objFolderItems, tmpFileList)

           '配列の大きさを再度求める
           lngCount = 0
           On Error Resume Next
           lngCount = UBound(tmpFileList) + 1
           On Error Goto 0

        Else
           'ファイルであれば、リストに格納します
           If Mid(objItem.Path, 2,1) = ":"  Or _
              Mid(objItem.Path, 2,2) = "\\" Then
              ReDim Preserve tmpFileList(lngCount)
              tmpFileList(lngCount) = objItem.Path
              lngCount = lngCount + 1
           End If

        End If

    Next

    'オブジェクトの破棄
    Set objItem = Nothing
    Set objFolderItems = Nothing

End Sub

'==========================================================
'= カラのzipフォルダを作成
'==========================================================
Function CreateZipMaster(ByVal pDir, ByVal pZipFilename)

    Dim objStream
    Dim objFS

    CreateZipMaster = False

    Set objFS = CreateObject("Scripting.FilesystemObject")
    Set objStream = Createobject("ADODB.Stream")

    '同名のzipが無い場合のみ作成
    If objFS.FileExists(objFS.BuildPath(pDir, pZipFilename)) = False Then

       'Zip作成
       With objStream
           .Open
           .Type = adTypeText
           .charset = "iso-8859-1"
           .WriteText ChrW(&h50) & ChrW(&h4B) & ChrW(&h05) & ChrW(&h06) _
                    & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) _
                    & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) _
                    & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) _
                    & ChrW(&h00) & ChrW(&h00) & ChrW(&h00)
           .SaveToFile objFS.BuildPath(pDir, pZipFilename) ,2
           .Close
       End With

       CreateZipMaster = True

    End If

    Set objStream = Nothing
    Set objFS = Nothing

End Function
