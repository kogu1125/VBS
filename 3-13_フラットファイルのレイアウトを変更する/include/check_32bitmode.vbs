Option Explicit

'64�r�b�g�N����e��
If UCase(CreateObject("WScript.Shell").Environment("Process").Item("PROCESSOR_ARCHITECTURE")) <> "X86" Then
   WScript.Echo "���̃v���O�����́A32�r�b�g���[�h�ŋN�����Ă��������B"
   WScript.Echo "���N����:"
   WScript.Echo "  C:\Windows\SysWOW64\CScript " & WScript.ScriptName & " [�����E�X�C�b�`�Ȃ�]"
   WScript.Quit(-1)
End If
