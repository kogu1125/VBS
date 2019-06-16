Option Explicit

'==========================================================
'= ファイルの有無チェック
'==========================================================
Function FileCheck(ByVal tmpPath, ByRef strPath)

    Dim objFS
    Dim strCheckPath
    Dim strCheckPath2

    'プロシージャの結果を初期化します
    FileCheck = False

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'チェックするフォルダのパス情報を作成（絶対パスに編集）
    strCheckPath = objFS.GetAbsolutePathName(tmpPath)

    '編集したパスでフォルダの有無をチェック
    If objFS.FileExists(strCheckPath) = True Then
       FileCheck = True
       strPath = strCheckPath 'パス情報を返す
    End If

    Set objFS = Nothing 'オブジェクトの破棄

End Function

'==========================================================
'= テキストファイルへの出力（ファイルの最後へ追記）
'==========================================================
Function TextWriteBottom(ByRef strOutfile, ByRef strMsg)

    Dim objFS
    Dim objOutfile

    TextWriteBottom = False

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    If objFS.FileExists(strOutfile) = False Then
       objFS.CreateTextFile(strOutfile)
    End If

    Set objOutfile = objFS.OpenTextfile(strOutfile, ForAppending)
    objOutfile.WriteLine strMsg
    objOutfile.Close

    TextWriteBottom = True

    Set objOutfile = Nothing
    Set objFS = Nothing

End Function

'==========================================================
'= テキストファイルへの出力（ファイルの先頭へ追記）
'==========================================================
Function TextWriteTop(ByRef strOutfile, ByRef strMsg)

    Dim objFS
    Dim objIOFile
    Dim strReadAll

    TextWriteTop = False

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    If objFS.FileExists(strOutfile) = False Then
       objFS.CreateTextFile(strOutfile)
    End If

    Set objIOFile = objFS.OpenTextfile(strOutfile, ForReading)
    If objIOFile.AtEndOfStream = False Then
       strReadAll = objIOFile.ReadAll
    Else
       strReadAll = vbNullString
    End If
    objIOFile.Close

    Set objIOFile = objFS.OpenTextfile(strOutfile, ForWriting)
    objIOFile.Write strMsg & vbCrLf & strReadAll
    objIOFile.Close

    TextWriteTop = True

    Set objIOFile = Nothing
    Set objFS = Nothing

End Function
