Option Explicit

'FilesystemObject�p
Const ForReading = 1   '�ǂݍ���
Const ForWriting = 2   '�������݁i�㏑�����[�h�j
Const ForAppending = 8 '�������݁i�ǋL���[�h�j

'WScript.Shell/Run�p
Const vbHide = 0
Const vbNormalFocus = 1
Const vbMinimizedFocus = 2
Const vbMaximizedFocus = 3
Const vbNormalNoFocus = 4
Const vbMinimizedNoFocus = 6

'BASP21 StrConv�p
Const vbUpperCase = 1
Const vbLowerCase = 2
Const vbProperCase = 3
Const vbWide = 4
Const vbNarrow = 8
Const vbKatakana = 16
Const vbHiragana = 32

'CopyHere MoveHere�p
Const FOF_SILENT            = &H04
Const FOF_RENAMEONCOLLISION = &H08
Const FOF_NOCONFIRMATION    = &H10
Const FOF_ALLOWUNDO         = &H40
Const FOF_FILESONLY         = &H80
Const FOF_SIMPLEPROGRESS    = &H100
Const FOF_NOCONFIRMMKDIR    = &H200
Const FOF_NOERRORUI         = &H400
Const FOF_NORECURSION       = &H1000
