Option Explicit

'FilesystemObject用
Const ForReading = 1   '読み込み
Const ForWriting = 2   '書きこみ（上書きモード）
Const ForAppending = 8 '書きこみ（追記モード）

'WScript.Shell/Run用
Const vbHide = 0
Const vbNormalFocus = 1
Const vbMinimizedFocus = 2
Const vbMaximizedFocus = 3
Const vbNormalNoFocus = 4
Const vbMinimizedNoFocus = 6

'BASP21 StrConv用
Const vbUpperCase = 1
Const vbLowerCase = 2
Const vbProperCase = 3
Const vbWide = 4
Const vbNarrow = 8
Const vbKatakana = 16
Const vbHiragana = 32

'CopyHere MoveHere用
Const FOF_SILENT            = &H04
Const FOF_RENAMEONCOLLISION = &H08
Const FOF_NOCONFIRMATION    = &H10
Const FOF_ALLOWUNDO         = &H40
Const FOF_FILESONLY         = &H80
Const FOF_SIMPLEPROGRESS    = &H100
Const FOF_NOCONFIRMMKDIR    = &H200
Const FOF_NOERRORUI         = &H400
Const FOF_NORECURSION       = &H1000
