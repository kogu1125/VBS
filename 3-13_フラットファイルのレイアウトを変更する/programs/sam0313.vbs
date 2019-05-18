Option Explicit

Function FlatToFlat()

    Dim objFS
    Dim objCon
    Dim strInFile
    Dim strInFilename
    Dim strInPath
    Dim strOutFile
    Dim strOutFilename
    Dim objOutFile
    Dim objRS
    Dim strSQL

    '�v���V�[�W���̌��ʂ����������܂�
    FlatToFlat = False

    '�R�}���h���C�������œ��́E�o�̓t�@�C�����w�肳��Ă��邩��
    '�`�F�b�N���܂�
    If WScript.Arguments.Count >= 2  Then

       '���̗͂L�����`�F�b�N���܂�
       If FileCheck(WScript.Arguments(0), strInFile) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : ���̓t�@�C�������݂��܂���B " & WScript.Arguments(0)
          Exit Function
       End If

       '�o�̗͂L�����`�F�b�N���܂�
       If FileCheck(WScript.Arguments(1), strOutFile) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �o�̓t�@�C�������݂��܂���B " & WScript.Arguments(1)
          Exit Function
       End If

    Else
       '�p�����^���������w�肳��Ă��Ȃ��Ƃ��́A�G���[���b�Z�[�W��\����
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �p�����^�𐳂����w�肵�Ă��������B"
       Exit Function
    End If

    'ADO�R�l�N�V�������쐬���܂�
    strInPath = getFilePath(strInFile)
    If getTextConnection(objCon, strInPath) = False Then
       '�R�l�N�V�������쐬�ł��Ȃ������Ƃ��̓v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �R�l�N�V�������쐬�ł��܂���ł����B"
       Exit Function
    End if

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '�o�̓t�@�C�����N���A���܂�
    Set objOutFile = objFS.OpenTextFile(strOutFile, ForWriting)
    objOutFile.Close

    '���́E�o�̓t�@�C�����̎擾
    strInFilename = objFS.GetFilename(strInFile)
    strOutFilename = objFS.GetFilename(strOutFile)

    '���̓t�@�C���̓ǂݍ���
    Set objRS = objCon.Execute("select * from " & strInFilename)

    Do Until objRS.Eof = True

       '�������ݗpSQL�̕ҏW
       strSQL = vbNullString
       strSQL = strSQL & "INSERT  INTO FLATOUTDATA.txt"
       strSQL = strSQL & "       (���t"
       strSQL = strSQL & "       ,�S����"
       strSQL = strSQL & "       ,���i�R�[�h"
       strSQL = strSQL & "       ,���i ) "
       strSQL = strSQL & "VALUES ('" & objRS("���t") & "'"
       strSQL = strSQL & "       ,'" & objRS("�S����") & "'"
       strSQL = strSQL & "       ,'" & objRS("���i�R�[�h") & "'"
       strSQL = strSQL & "       ,'" & Right("0000" & objRS("���i"), 5) & "' )"

       '�o�̓t�@�C���֏�������
       objCon.Execute(strSQL)

       '���̃��R�[�h��ǂ�
       objRS.MoveNext

    Loop

    '�R�l�N�V�����ƃ��R�[�h�Z�b�g�̃N���[�Y
    objRS.Close
    objCON.Close

    WScript.echo "�������������܂���"

    '�I�u�W�F�N�g�̔j��
    Set objRS = Nothing
    Set objCon = Nothing
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    FlatToFlat = True

End Function
