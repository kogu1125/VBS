Option Explicit

'64ビット起動を弾く
If UCase(CreateObject("WScript.Shell").Environment("Process").Item("PROCESSOR_ARCHITECTURE")) <> "X86" Then
   WScript.Echo "このプログラムは、32ビットモードで起動してください。"
   WScript.Echo "◆起動例:"
   WScript.Echo "  C:\Windows\SysWOW64\CScript " & WScript.ScriptName & " [引数・スイッチなど]"
   WScript.Quit(-1)
End If
