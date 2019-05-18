Option Explicit

Function FileCutter()

    Dim objWSArg
    Dim objFS
    Dim objADOST_R
    Dim objADOST_W
    Dim strTargetFile
    Dim strCreatePath
    Dim strCreateName
    Dim strCreateFilename
    Dim lngDivCount
    Dim lngDivSerial
    Dim bytData

    Dim lngDivParm
    Dim lngDivSize

    '�v���V�[�W���̌��ʂ����������܂�
    FileCutter = False

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'objWSArg�I�u�W�F�N�g�̃C���X�^���X���R�s�[
    Set objWSArg = WScript.Arguments

    '�R�}���h���C�������őΏۂ̃t�@�C�����w�肳��Ă��邩��
    '�`�F�b�N���܂�
    If objWSArg.Unnamed.Count = 1 Then

       '�Ώۃt�@�C���̗L�����`�F�b�N���܂��B
       If FileCheck(objWSArg.Unnamed(0), strTargetFile) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �Ώۃt�@�C�������݂��܂���B " & objWSArg.Unnamed(0)
          Exit Function
       End If

    Else
       '�p�����^���������w�肳��Ă��Ȃ��Ƃ��́A�G���[���b�Z�[�W��\����
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �p�����^�𐳂����w�肵�Ă��������B"
       Exit Function
    End If

    '�������̃`�F�b�N
    If objWSArg.Named.Count = 0 Then
       '�������ȗ�����2��
       lngDivParm = 2

    '�X�C�b�`������̂ɕ������������ꍇ
    ElseIf objWSArg.Named.Exists("c") = False Then
       WScript.Echo "ERROR : �������� /c:?? �Ŏw�肵�Ă��������B"
       Exit Function
    Else

       '�������ȗ�����2��
       lngDivParm = objWSArg.Named("c")

       '�����������l�ȊO�̏ꍇ�̓G���[
       If IsNumeric(lngDivParm) = False Then
          WScript.Echo "ERROR : �������͐��l�Ŏw�肵�Ă��������B"
          Exit Function
       ElseIf CLng(lngDivParm) < 2 Then
          WScript.Echo "ERROR : ��������2�ȏ�Ŏw�肵�Ă��������B"
          Exit Function
       End If

       '�O�̂��ߐ��l�^�ɕϊ�
       lngDivParm = CLng(lngDivParm)

    End If

    '������̑傫�������߂�(�[���͐؂�グ)
    lngDivSize = Fix(objFS.GetFile(strTargetFile).Size / lngDivParm + 0.99999999)

    '���������t�@�C���̃o�C�g�����������Ƃ��͕����ł��Ȃ�
    If objFS.GetFile(strTargetFile).Size < lngDivParm Then
       Set objFS = Nothing
       WScript.Echo "ERROR : " & lngDivParm & "�ɂ͕����ł��܂���B"
       Exit Function
    End if

    'ADODB.Stream�I�u�W�F�N�g���쐬���܂�
    Set objADOST_W = Createobject("ADODB.Stream") 'READ�p
    Set objADOST_R = Createobject("ADODB.Stream") 'WRITE�p

    '�������̃J�E���g�p
    lngDivSerial = 0
    lngDivCount  = 0

    strCreatePath = getFilePath(strTargetFile)

    '���̓t�@�C����ADODB.Stream�ŃI�[�v�����܂�
    objADOST_R.Type = adTypeBinary
    objADOST_R.Open
    objADOST_R.LoadFromFile strTargetFile
    objADOST_R.Position = 0

    '�t�@�C���̕�������
    Do Until objADOST_R.EOS = True

       '�������̃J�E���g
       lngDivSerial = lngDivSerial + 1
       lngDivCount  = lngDivCount + 1

       '[POINT!]������T�C�Y����ǂݍ��݂܂�
       bytData = objADOST_R.Read(lngDivSize)

       '�ǂݍ��񂾃f�[�^�𕪊���t�@�C���֏o�͂��܂�
       objADOST_W.Open
       objADOST_W.Type = adTypeBinary
       objADOST_W.Write bytData

       '[POINT!]�t�@�C�����ɂ͕����ԍ���t�����܂�
       strCreateName = objFS.GetFilename(strTargetFile) & _
                       "." & Right("000" & lngDivSerial, 3) & ".div"
       strCreateFilename = objFS.BuildPath(strCreatePath, strCreateName)

       WScript.echo "���� -> " & strCreateFilename

       '�O�̂��ߏo�̓t�@�C�����폜���܂�
       On Error Resume Next
       objFS.DeleteFile strCreateFilename
       On Error GoTo 0

       '�t�@�C�����o�͂��܂�
       objADOST_W.SaveToFile strCreateFilename

       objADOST_W.Close

    Loop

    '���̓t�@�C�����N���[�Y
    objADOST_R.Close

    '�������ʂ̕\��
    If lngDivCount > 1 Then
       WScript.Echo lngDivCount & "�̃t�@�C���ɕ������܂����B"
    End If

    '�I�u�W�F�N�g�̔j��
    Set objADOST_R = Nothing
    Set objADOST_W = Nothing
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    FileCutter = True

End Function
