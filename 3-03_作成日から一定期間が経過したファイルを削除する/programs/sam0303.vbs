Option Explicit

Function DeleteTimelimit()

    Dim objFS
    Dim strExecFolder
    Dim lngDays
    Dim lngFileDays
    Dim strFileName
    Dim strList
    Dim lngDelCount
    Dim blnSubFolder
    Dim lngSwitchCount

    '�v���V�[�W���̌��ʂ����������܂�
    DeleteTimelimit = False

    '�R�}���h���C�������ŃR�s�[���t�@�C�����ƃR�s�[��t�H���_����
    '�w�肳��Ă��邩���`�F�b�N���܂�
    If WScript.Arguments.Unnamed.Count = 2 Then

       '�����t�H���_�̗L�����`�F�b�N���܂�
       If FolderCheck(WScript.Arguments.Unnamed(0), strExecFolder) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\����
          '�v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �w�肵���t�H���_�͑��݂��܂���B " & WScript.Arguments.Unnamed(0)
          Exit Function
       End If

       '�����p�����^���擾
       lngDays = WScript.Arguments.Unnamed(1)

       '���������l���ǂ������`�F�b�N���܂�
       If IsNumeric(lngDays) = True Then

          lngDays = CLng(lngDays) '�O�̂��ߐ����^�ɕϊ����܂�

          '������0�����̂Ƃ��̓p�����^�G���[�Ńv���V�[�W�����I��
          If lngDays < 0 Then
             WScript.Echo "ERROR : �����ɂ�0�ȏ���w�肵�Ă��������B"
             Exit Function
          End If
       Else
          WScript.Echo "ERROR : �����͐��l�Ŏw�肵�Ă��������B"
          Exit Function
       End If

    Else
       '�p�����^���������w�肳��Ă��Ȃ��Ƃ��́A�G���[���b�Z�[�W��\����
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �p�����^�𐳂����w�肵�Ă��������B"
       Exit Function
    End If

    '�X�C�b�`�����擾
    lngSwitchCount = WScript.Arguments.Named.Count

    '/sub�X�C�b�`�̃`�F�b�N
    If WScript.Arguments.Named.Exists("sub") = False Then
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

    lngDelCount = 0 '�폜�t�@�C���̃J�E���g�p

    '�t�@�C�����X�g�̎擾
    If blnSubFolder = False Then
       strList = SearchFolder(strExecFolder)     '/sub���w�莞
    Else
       strList = sSearchFolderAll(strExecFolder) '/sub�w�莞
    End If

    '[POINT!]�t�@�C�����X�g�̏���
    For Each strFileName In strList

        '[POINT!]�t�@�C���쐬������̌o�ߓ��������߂�
        On Error Resume Next 'ZIP��LZH�̂Ƃ��Ɏ��s���G���[�𔭐������Ȃ��悤�ɂ���
        lngFileDays = -1
        lngFileDays = DateDiff("d", objFS.GetFile(strFileName).DateCreated, Date)
        On Error GoTo 0

        '[POINT!]�w�肵�����Ԃ��o�߂��Ă���΁A���̃t�@�C�����폜���܂�
        If lngFileDays >= lngDays Then
           objFS.DeleteFile strFileName
           WScript.Echo "�폜���܂��� " & strFileName
           lngDelCount = lngDelCount + 1
        End If

    Next

    '�������ʃ��b�Z�[�W�̕\��
    If lngDelCount > 0 Then
       WScript.Echo lngDelCount & "�̃t�@�C�����폜���܂���"
    Else
       WScript.Echo "�폜�ł���t�@�C���͂���܂���ł����B"
    End If

    '�I�u�W�F�N�g�̔j��
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    DeleteTimelimit = True

End Function
