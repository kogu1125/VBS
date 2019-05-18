Option Explicit

'==========================================================
'= プログラムが保存されているパスを取得
'==========================================================
Function getScriptDir()

    getScriptDir = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

End Function

'==========================================================
'= カレントディレクトリの取得
'==========================================================
Function getCurrentDir()

    Dim objWSh

    'カレントディレクトリを取得します
    Set objWSh = CreateObject("WScript.Shell")
    getCurrentDir = objWSh.CurrentDirectory
    Set objWSh = Nothing

End Function

'==========================================================
'= 偶数・奇数を評価(True - 偶数, False - 奇数)
'==========================================================
Function IsEven(tmpNum)

    If (tmpNum Mod 2) = 0 Then
       IsEven = True
    Else
       IsEven = False
    End If

End Function

'==========================================================
'= ファイル名を含むフルパスから、パス情報のみを取得
'==========================================================
Function getFilePath(tmpPath)

    Dim objFS
    Dim strFilename

    'ファイルシステムオブジェクトを作成
    Set objFS = Createobject("Scripting.FilesystemObject")

    'ファイル名を含むパス情報から、ファイル名を含まないパス情報を作成
    strFilename = objFS.GetFilename(tmpPath)
    getFilePath = Replace(tmpPath, strFilename, vbNullString)

    'オブジェクトを破棄
    Set objFS = Nothing

End Function

'==========================================================
'= コマンドライン上でプログレスバーを表示する
'==========================================================
Sub cmdProgressBar(Total, Count)

    Dim lngPercent
    Dim lngCount
    Dim i

    '現在の割合を計算
    lngPercent = Int(Count / Total * 100)
    '表示カウンタ
    lngCount = lngPercent / 2

    WScript.StdOut.Write vbCr
    WScript.StdOut.Write Right("  " & lngPercent,3) & "% "

    '棒グラフの表示
    For i=1 to lngCount
        WScript.StdOut.Write "|"
    Next

End Sub

'==========================================================
'= ランダムな文字列を返す
'==========================================================
Function getRandomString(ByVal tmpLength)

    Const cChrs = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

    Dim strRString
    Dim lngLen
    Dim lngRnd
    Dim i

    Randomize

    '基本文字セットの長さを求める
    lngLen = Len(cChrs)

    strRString = vbNullString

    '依頼長だけ繰り返す
    For i=1 To tmpLength

        '取り出す位置は乱数で求める
        lngRnd = Int(Rnd * lngLen) + 1

        '基本文字セットから１文字取り出し、strRStringへ加える
        strRString = strRString & Mid(cChrs, lngRnd, 1)

    Next

    'ランダムな文字列を戻り値へ返す
    getRandomString = strRString

End Function
