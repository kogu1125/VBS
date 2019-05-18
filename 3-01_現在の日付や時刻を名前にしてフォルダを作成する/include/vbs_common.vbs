Option Explicit

'==========================================================
'= �v���O�������ۑ�����Ă���p�X���擾
'==========================================================
Function getScriptDir()

    getScriptDir = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

End Function

'==========================================================
'= �J�����g�f�B���N�g���̎擾
'==========================================================
Function getCurrentDir()

    Dim objWSh

    '�J�����g�f�B���N�g�����擾���܂�
    Set objWSh = CreateObject("WScript.Shell")
    getCurrentDir = objWSh.CurrentDirectory
    Set objWSh = Nothing

End Function

'==========================================================
'= �����E���]��(True - ����, False - �)
'==========================================================
Function IsEven(tmpNum)

    If (tmpNum Mod 2) = 0 Then
       IsEven = True
    Else
       IsEven = False
    End If

End Function

'==========================================================
'= �t�@�C�������܂ރt���p�X����A�p�X���݂̂��擾
'==========================================================
Function getFilePath(tmpPath)

    Dim objFS
    Dim strFilename

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬
    Set objFS = Createobject("Scripting.FilesystemObject")

    '�t�@�C�������܂ރp�X��񂩂�A�t�@�C�������܂܂Ȃ��p�X�����쐬
    strFilename = objFS.GetFilename(tmpPath)
    getFilePath = Replace(tmpPath, strFilename, vbNullString)

    '�I�u�W�F�N�g��j��
    Set objFS = Nothing

End Function

'==========================================================
'= �R�}���h���C����Ńv���O���X�o�[��\������
'==========================================================
Sub cmdProgressBar(Total, Count)

    Dim lngPercent
    Dim lngCount
    Dim i

    '���݂̊������v�Z
    lngPercent = Int(Count / Total * 100)
    '�\���J�E���^
    lngCount = lngPercent / 2

    WScript.StdOut.Write vbCr
    WScript.StdOut.Write Right("  " & lngPercent,3) & "% "

    '�_�O���t�̕\��
    For i=1 to lngCount
        WScript.StdOut.Write "|"
    Next

End Sub

'==========================================================
'= �����_���ȕ������Ԃ�
'==========================================================
Function getRandomString(ByVal tmpLength)

    Const cChrs = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

    Dim strRString
    Dim lngLen
    Dim lngRnd
    Dim i

    Randomize

    '��{�����Z�b�g�̒��������߂�
    lngLen = Len(cChrs)

    strRString = vbNullString

    '�˗��������J��Ԃ�
    For i=1 To tmpLength

        '���o���ʒu�͗����ŋ��߂�
        lngRnd = Int(Rnd * lngLen) + 1

        '��{�����Z�b�g����P�������o���AstrRString�։�����
        strRString = strRString & Mid(cChrs, lngRnd, 1)

    Next

    '�����_���ȕ������߂�l�֕Ԃ�
    getRandomString = strRString

End Function
