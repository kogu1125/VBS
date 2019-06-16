Option Explicit

Function TextReplace()

    Dim objFS
    Dim strTargetFile
    Dim strBefore()
    Dim strAfter()
    Dim i
    Dim j
    Dim objInFile
    Dim strData
    Dim strCtrlA
    Dim strCtrlB
    Dim objNewFile
    Dim strNewFilePath
    Dim strNewFile

    '�v���V�[�W���̌��ʂ����������܂�
    TextReplace = False

    '�R�}���h���C�������̐�3�ȏ�A������ǂ������`�F�b�N���܂��B
    If WScript.Arguments.Count >= 3 And _
       IsEven(WScript.Arguments.Count) = False Then

       '�^�[�Q�b�g�t�@�C���̗L�����`�F�b�N���܂�
       If FileCheck(WScript.Arguments(0), strTargetFile) = False Then
          '�p�X��������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �^�[�Q�b�g�t�@�C�������݂��܂���B " & WScript.Arguments(0)
          Exit Function
       End If

       j = 0

       '�u���������̕ۑ�
       For i=2 To WScript.Arguments.Count Step 2
           ReDim Preserve strBefore(j)
           ReDim Preserve strAfter(j)
           strBefore(j) = WScript.Arguments(i -1)
           strAfter(j) = WScript.Arguments(i)
           j = j + 1
       Next

    Else
       '�p�����^���������w�肳��Ă��Ȃ��Ƃ��́A�G���[���b�Z�[�W��\����
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �p�����^�𐳂����w�肵�Ă��������B"
       Exit Function
    End If

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '���t�@�C����ǂݍ��݂܂�
    Set objInFile = objFS.OpenTextFile(strTargetFile, ForReading)
    strData = objInFile.ReadAll
    objInFile.Close

    '�u�������������s���܂�(For�I���l�͔z��̑傫��)
    For i=0 To UBound(strBefore)

        '�u�������O������̕ҏW
        strCtrlB = strBefore(i)
        If UCase(strCtrlB) = "/S" Then       '���䕶��(���p��)
           strCtrlB = " "
        ElseIf UCase(strCtrlB) = "/W" Then   '���䕶��(�S�p��)
           strCtrlB = "�@"
        End If

        '�u�������㕶����̕ҏW
        strCtrlA = strAfter(i)
        If UCase(strCtrlA) = "/S" Then       '���䕶��(���p��)
           strCtrlA = " "
        ElseIf UCase(strCtrlA) = "/W" Then   '���䕶��(�S�p��)
           strCtrlA = "�@"
        ElseIf UCase(strCtrlA) = "/NUL" Then '���䕶��(�J������)
           strCtrlA = vbNullString
        End If

        '�u����������
        strData = Replace(strData, strCtrlB, strCtrlA)

    Next

    '���ԃt�@�C���̍쐬
    strNewFilePath = getFilePath(strTargetFile)
    strNewFile = objFS.BuildPath(strNewFilePath, objFS.GetTempName)
    Set objNewFile = objFS.CreateTextFile(strNewFile, ForWriting)

    '���ԃt�@�C���ւ̏�������
    objNewFile.Write strData
    objNewFile.Close

    '���ԃt�@�C�����^�[�Q�b�g�t�@�C���ɂ���
    objFS.DeleteFile strTargetFile
    objFS.MoveFile strNewFile, strTargetFile

    WScript.echo "�u���������������܂��� " & strTargetFile

    '�I�u�W�F�N�g�̔j��
    Set objInFile = Nothing
    Set objNewFile = Nothing
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    TextReplace = True

End Function
