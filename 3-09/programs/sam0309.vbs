Option Explicit

Function FileJoin()

    Dim objWSArg
    Dim objFS
    Dim objADOST_R
    Dim objADOST_W
    Dim strCopyFrom
    Dim strCopyTo
    Dim strCopyFilename
    Dim strCopyFileBasename
    Dim strCopyFileExt
    Dim strFileName
    Dim strList
    Dim lngCopyCount
    Dim strTempFilename
    Dim bytData
    Dim strJoinExt
    Dim blnSubFolder
    Dim lngSwitchCount

    '�v���V�[�W���̌��ʂ����������܂�
    FileJoin = False

    'objWSArg�I�u�W�F�N�g�̃C���X�^���X���R�s�[
    Set objWSArg = WScript.Arguments

    '�R�}���h���C�������ŃR�s�[���t�H���_���ƃR�s�[��t�@�C������
    '�w�肳��Ă��邩���`�F�b�N���܂�
    If objWSArg.Unnamed.Count = 2 Then

       '�R�s�[���t�H���_�̗L�����`�F�b�N���܂�
       If FolderCheck(objWSArg.Unnamed(0), strCopyFrom) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �R�s�[���t�H���_�͑��݂��܂���B " & objWSArg.Unnamed(0)
          Exit Function
       End If

       '�R�s�[��t�H���_�̗L�����`�F�b�N���܂�
       '�������A�˗��p�X���ɂ̓t�@�C�������܂݂܂�
       If FileFolderCheck(objWSArg.Unnamed(1), strCopyTo) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �쐬��̃p�X�����݂��܂���B " & objWSArg.Unnamed(1)
          Exit Function
       End If

    Else
       '�p�����^���������w�肳��Ă��Ȃ��Ƃ��́A�G���[���b�Z�[�W��\����
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �p�����^�𐳂����w�肵�Ă��������B"
       Exit Function
    End If

    lngSwitchCount = WScript.Arguments.Named.Count

    '�g���q�w��̈���
    If objWSArg.Named.Exists("e") = False Then
       WScript.Echo "ERROR : ��������t�@�C���̊g���q�� /e:?? �Ŏw�肵�Ă��������B"
       Exit Function
    Else
       '�����t�@�C���̊g���q���擾
       strJoinExt = objWSArg.Named("e")

       '�K�{���̓`�F�b�N
       If strJoinExt = vbNullString Then
          WScript.Echo "ERROR : �g���q���w�肳��Ă��܂���B"
          Exit Function
       End If

       lngSwitchCount = lngSwitchCount - 1

    End If

    '/sub�X�C�b�`�̃`�F�b�N
    If objWSArg.Named.Exists("sub") = False Then
       blnSubFolder = False
    Else
       blnSubFolder = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '�X�C�b�`��ނ̃`�F�b�N
    If lngSwitchCount > 0 Then '�]���ȃX�C�b�`������
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �����ȃX�C�b�`������܂��B"
       Exit Function
    End If

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")
    'ADODB.Stream�I�u�W�F�N�g���쐬���܂�
    Set objADOST_W = Createobject("ADODB.Stream") 'READ�p
    Set objADOST_R = Createobject("ADODB.Stream") 'WRITE�p

    '�R�s�[�t�@�C���̃J�E���g�p
    lngCopyCount = 0

    '�t�@�C�����X�g�̎擾
    If blnSubFolder = False Then
       strList = SearchFolder(strCopyFrom)     '/sub���w�莞
    Else
       strList = sSearchFolderAll(strCopyFrom) '/sub�w�莞
    End If

    '���ԃt�@�C�����쐬���܂�
    strTempFilename = objFS.BuildPath(strCopyFrom, objFS.GetTempName)

    '�o�̓t�@�C����ADODB.Stream(�o�C�i�����[�h)�ŊJ���܂�
    objADOST_W.Open
    objADOST_W.Type = adTypeBinary

    '[POINT!]�t�@�C�����X�g�̏���
    For Each strFileName In strList

        '�t���p�X����t�@�C�������擾���܂�
        strCopyFilename = objFS.GetFilename(strFileName)
        '�t���p�X����t�@�C�������擾���܂�
        strCopyFileBasename = objFS.GetBasename(strFileName)
        '�������A�g���q���擾���܂�
        strCopyFileExt  = objFS.GetExtensionName(strFileName)

        '[POINT!]�g���q����v���邩�ǂ����𔻒f���܂��i�啶���Ŕ��f�j
        If UCase(strCopyFileExt) = UCase(strJoinExt) Then

           '�t�@�C��ADODB.Stream�œǂݍ��݂܂�
           objADOST_R.Type = adTypeBinary
           objADOST_R.Open
           objADOST_R.LoadFromFile strFileName
           objADOST_R.Position = 0
           bytData = objADOST_R.Read()
           objADOST_R.Close

           '�t�@�C�����ǂݍ��߂��Ƃ��͒��ԃt�@�C���֏o�͂��鏀�������܂�
           If IsNull(bytData)=False Then
              objADOST_W.Write bytData '�o�b�t�@�֏o��
           End If

           WScript.Echo "�������܂� : " & strFileName
           lngCopyCount = lngCopyCount + 1
        End If

    Next

    '�����t�@�C���̍쐬�ƁA�������ʃ��b�Z�[�W�̕\��
    If lngCopyCount > 0 Then

       WScript.Echo lngCopyCount & "�̃t�@�C�����������܂����B"

       '�o�̓t�@�C����ۑ����܂�
       objADOST_W.SaveToFile strTempFilename
       objADOST_W.Close

       '���ԃt�@�C�����쐬��ɃR�s�[���A���ԃt�@�C�����폜���܂�
       objFS.CopyFile strTempFilename, strCopyTo
       objFS.DeleteFile strTempFilename

    Else
       WScript.Echo "��������t�@�C���͂���܂���ł����B"
    End If

    '�I�u�W�F�N�g�̔j��
    Set objADOST_R = Nothing
    Set objADOST_W = Nothing
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    FileJoin = True

End Function
