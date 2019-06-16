Option Explicit

'==========================================================
'= �t�@�C���̗L���`�F�b�N
'==========================================================
Function FileCheck(ByVal tmpPath, ByRef strPath)

    Dim objFS
    Dim strCheckPath
    Dim strCheckPath2

    '�v���V�[�W���̌��ʂ����������܂�
    FileCheck = False

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '�`�F�b�N����t�H���_�̃p�X�����쐬�i��΃p�X�ɕҏW�j
    strCheckPath = objFS.GetAbsolutePathName(tmpPath)

    '�ҏW�����p�X�Ńt�H���_�̗L�����`�F�b�N
    If objFS.FileExists(strCheckPath) = True Then
       FileCheck = True
       strPath = strCheckPath '�p�X����Ԃ�
    End If

    Set objFS = Nothing '�I�u�W�F�N�g�̔j��

End Function

'==========================================================
'= �e�L�X�g�t�@�C���ւ̏o�́i�t�@�C���̍Ō�֒ǋL�j
'==========================================================
Function TextWriteBottom(ByRef strOutfile, ByRef strMsg)

    Dim objFS
    Dim objOutfile

    TextWriteBottom = False

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
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
'= �e�L�X�g�t�@�C���ւ̏o�́i�t�@�C���̐擪�֒ǋL�j
'==========================================================
Function TextWriteTop(ByRef strOutfile, ByRef strMsg)

    Dim objFS
    Dim objIOFile
    Dim strReadAll

    TextWriteTop = False

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
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
