Option Explicit

'コンソール起動を弾く
If UCase(Right(WScript.FullName, 11)) = "CSCRIPT.EXE" Then
   WScript.Echo "ERROR : このプログラムは、GUIモードで起動してください。"
   WScript.Quit(-1)
End If
