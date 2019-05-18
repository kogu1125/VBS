Option Explicit

'GUI起動を弾く
If UCase(Right(WScript.FullName, 11)) = "WSCRIPT.EXE" Then
   MsgBox "このプログラムは、コマンドプロンプトで起動してください。", vbOkOnly + vbExclamation, "起動モードエラー"
   WScript.Quit(-1)
End If
