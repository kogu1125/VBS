Option Explicit

Function CopySelectType()

    Dim objWSArg
    Dim objFS
    Dim strCopyFrom
    Dim strCopyFilename
    Dim strCopyFileExt
    Dim strFileName
    Dim strList
    Dim lngCopyCount
    Dim strExtName
    Dim blnSubFolder
    Dim lngSwitchCount
    Dim fn 

    '�v���V�[�W���̌��ʂ����������܂�
    CopySelectType = False

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'objWSArg�I�u�W�F�N�g�̃C���X�^���X���R�s�[
    Set objWSArg = WScript.Arguments

    '�R�}���h���C�������ŃR�s�[���t�@�C�����ƃR�s�[��t�H���_����
    '�w�肳��Ă��邩���`�F�b�N���܂�
    If objWSArg.Unnamed.Count = 1 Then

       '�������t�H���_�̗L�����`�F�b�N���܂�
       If FolderCheck(objWSArg.Unnamed(0), strCopyFrom) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �������t�H���_�͑��݂��܂���B " & objWSArg.Unnamed(0)
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
       WScript.Echo "ERROR : �g���q�� /e:??? �Ŏw�肵�Ă��������B "
       Exit Function
    Else
       '���o�����i�g���q�j���擾
       strExtName = objWSArg.Named("e")

       '�K�{���̓`�F�b�N
       If strExtName = vbNullString Then
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

    '�R�s�[�t�@�C���̃J�E���g�p
    lngCopyCount = 0

    '�t�@�C�����X�g�̎擾
    If blnSubFolder = False Then
       strList = SearchFolder(strCopyFrom)     '/sub���w�莞
    Else
       strList = sSearchFolderAll(strCopyFrom) '/sub�w�莞
    End If

    '[POINT!]�t�@�C�����X�g�̏���
    For Each strFileName In strList

        '�t�@�C�������擾���܂�
        strCopyFilename = objFS.GetFilename(strFileName)

        '�g���q���擾���܂�
        strCopyFileExt = objFS.GetExtensionName(strFileName)

        '[POINT!]�g���q����v���邩�ǂ����𔻒f���܂�
        '        ���̂Ƃ��ɑ啶���Ŕ��f���܂�
        If InStr(1,strCopyFilename,strExtName,vbTextCompare) > 0 Then
           '�R�s�[���t�@�C�������������܂�
           Set fn = objFS.GetFile(strFileName) 
           WScript.Echo "���O�F" & fn.Name & "------�X�V�����F" & fn.DateLastModified & "------�T�C�Y�F" & FormatNumber(fn.Size, 0)
           WScript.Echo "�p�X:" & strFileName

           lngCopyCount = lngCopyCount + 1
        End If

    Next

    '�������ʃ��b�Z�[�W�̕\��
    If lngCopyCount > 0 Then
       WScript.Echo lngCopyCount & "�̃t�@�C�����������܂����B"
    Else
       WScript.Echo "��������t�@�C���͂���܂���ł����B"
    End If

    '�I�u�W�F�N�g�̔j��
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    CopySelectType = True

End Function
