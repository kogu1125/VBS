Option Explicit

Function ModifiedFileCopy()

    Dim objFS
    Dim strCopyFrom
    Dim strCopyTo
    Dim lngCopyDays
    Dim blnCopyCreate
    Dim blnCopyUpdate
    Dim blnCopyAccess
    Dim dteTargetDate
    Dim strCopyFilename
    Dim strCopyToFilename
    Dim blnCopyLockOn
    Dim strFileName
    Dim strList
    Dim lngCopyCount
    Dim blnSubFolder
    Dim lngSwitchCount
    Dim dteDateCreated
    Dim dteDateLastModified
    Dim dteDateLastAccessed

    '�v���V�[�W���̌��ʂ����������܂�
    ModifiedFileCopy = False

    '�R�}���h���C�������ŃR�s�[���t�@�C�����ƃR�s�[��t�H���_����
    '�w�肳��Ă��邩���`�F�b�N���܂�
    If WScript.Arguments.Unnamed.Count = 2 Or _
       WScript.Arguments.Unnamed.Count = 3 Then

       '�R�s�[���t�H���_�̗L�����`�F�b�N���܂�
       If FolderCheck(WScript.Arguments.Unnamed(0), strCopyFrom) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �R�s�[���t�H���_�͑��݂��܂���B " & WScript.Arguments.Unnamed(0)
          Exit Function
       End If

       '�R�s�[��t�H���_�̗L�����`�F�b�N���܂�
       If FolderCheck(WScript.Arguments.Unnamed(1), strCopyTo) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �R�s�[��t�H���_�͑��݂��܂���B " & WScript.Arguments.Unnamed(1)
          Exit Function
       End If

       lngCopyDays = 0

       '�������`�F�b�N
       If WScript.Arguments.Unnamed.Count = 3 Then

          If IsNumeric(WScript.Arguments.Unnamed(2)) = True Then
             lngCopyDays = CLng(WScript.Arguments.Unnamed(2))

             If lngCopyDays < 0 Then
                '0�ȉ����w�肵���Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
                WScript.Echo "ERROR : ������0�ȏ�Ŏw�肵�Ă��������B " & WScript.Arguments.Unnamed(2)
                Exit Function
             End If

          End If

       End If

    Else
       '�p�����^���������w�肳��Ă��Ȃ��Ƃ��́A�G���[���b�Z�[�W��\����
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �p�����^�𐳂����w�肵�Ă��������B"
       Exit Function
    End If

    '�X�C�b�`�����擾
    lngSwitchCount = WScript.Arguments.Named.Count

    '/c /u /a�X�C�b�`�̃`�F�b�N
    ':�쐬���őI�ʃX�C�b�`
    If WScript.Arguments.Named.Exists("c") = False Then
       blnCopyCreate = False
    Else
       blnCopyCreate = True
       lngSwitchCount = lngSwitchCount - 1
    End If
    ':�X�V���őI�ʃX�C�b�`
    If WScript.Arguments.Named.Exists("u") = False Then
       blnCopyUpdate = False
    Else
       blnCopyUpdate = True
       lngSwitchCount = lngSwitchCount - 1
    End If
    ':�ŏI�A�N�Z�X���őI�ʃX�C�b�`
    If WScript.Arguments.Named.Exists("a") = False Then
       blnCopyAccess = False
    Else
       blnCopyAccess = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/sub�X�C�b�`�̃`�F�b�N
    If WScript.Arguments.Named.Exists("sub") = False Then
       blnSubFolder = False
    Else
       blnSubFolder = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/c /u /a�̂�������w�肳��Ă��Ȃ��Ƃ��͏]���d�l���w�肵�����Ƃɂ���
    If blnCopyCreate = False And _
       blnCopyUpdate = False And _
       blnCopyAccess = False Then

       blnCopyCreate = True
       blnCopyUpdate = True

    End If

    '�X�C�b�`��ނ̃`�F�b�N
    If lngSwitchCount > 0 Then '�]���ȃX�C�b�`������
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �����ȃX�C�b�`������܂��B"
       Exit Function
    End If

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '�R�s�[�t�@�C���̃J�E���g�p
    lngCopyCount = 0

    '�t�@�C�����X�g�̎擾
    If blnSubFolder = False Then
       strList = SearchFolder(strCopyFrom)     '/sub���w�莞
    Else
       strList = sSearchFolderAll(strCopyFrom) '/sub�w�莞
    End If

    dteTargetDate = CStr(DateAdd("d", (lngCopyDays * (-1)), Date))

    '[POINT!]�t�@�C�����X�g�̏���
    For Each strFileName In strList

        '[POINT!]�t�@�C���̃^�C���X�^���v���擾����
        On Error Resume Next 'ZIP��LZH�̂Ƃ��Ɏ��s���G���[�𔭐������Ȃ��悤�ɂ���
        dteDateCreated = vbNullString
        dteDateCreated = FormatDateTime(objFS.GetFile(strFileName).DateCreated, vbShortDate)
        dteDateLastModified = vbNullString
        dteDateLastModified = FormatDateTime(objFS.GetFile(strFileName).DateLastModified, vbShortDate)
        dteDateLastAccessed = vbNullString
        dteDateLastAccessed = FormatDateTime(objFS.GetFile(strFileName).DateLastAccessed, vbShortDate)
        On Error GoTo 0

        blnCopyLockOn = False

        '[POINT!]�Ώۂ̃t�@�C�����ǂ����𔻒f���܂��B
        If blnCopyCreate  = True          And _
           dteDateCreated = dteTargetDate Then
           blnCopyLockOn = True
        End If

        If blnCopyUpdate       = True          And _
           dteDateLastModified = dteTargetDate Then
           blnCopyLockOn = True
        End If

        If blnCopyAccess       = True          And _
           dteDateLastAccessed = dteTargetDate Then
           blnCopyLockOn = True
        End If

        '�R�s�[�Ώۃt�@�C���̂Ƃ�
        If blnCopyLockOn = True Then
           '�t���p�X����t�@�C�������擾���܂�
           strCopyFilename = objFS.GetFilename(strFileName)
           '�R�s�[��t�@�C������ҏW���܂�
           strCopyToFilename = objFS.BuildPath(strCopyTo, strCopyFilename)
           '�t�@�C�����R�s�[���܂�
           objFS.CopyFile strFileName, strCopyToFilename
           WScript.Echo "�R�s�[���܂��� " & strFileName
           lngCopyCount = lngCopyCount + 1
        End If

    Next

    '�������ʃ��b�Z�[�W�̕\��
    If lngCopyCount > 0 Then
       WScript.Echo lngCopyCount & "�̃t�@�C�����R�s�[���܂����B"
    Else
       WScript.Echo "�R�s�[����t�@�C���͂���܂���ł����B"
    End If

    '�I�u�W�F�N�g�̔j��
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    ModifiedFileCopy = True

End Function
