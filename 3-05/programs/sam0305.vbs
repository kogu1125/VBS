Option Explicit

Dim cDelimiter

'[POINT!]��؂蕶��
cDelimiter = vbTab

Function FolderSearch()

    Dim objFS
    Dim strExecFolder
    Dim varFilename
    Dim strPrintData
    Dim strList
    Dim blnSubFolder
    Dim lngSwitchCount
    Dim blnHeader
    Dim blnFileName
    Dim blnFileSize
    Dim blnFileCreate
    Dim blnFileUpdate
    Dim blnFileAccess
    Dim blnZipFile
    Dim varFileSize
    Dim dteDateCreated
    Dim dteDateLastModified
    Dim dteDateLastAccessed

    '�v���V�[�W���̌��ʂ����������܂�
    FolderSearch = False

    '�R�}���h���C�������Ɍ����Ώۃt�H���_���w�肳��Ă���Ƃ�
    If WScript.Arguments.Unnamed.Count = 1 Then

       '�����t�H���_�̗L�����`�F�b�N���܂�
       If FolderCheck(WScript.Arguments.Unnamed(0), strExecFolder) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\����
          '�v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �w�肵���t�H���_�͑��݂��܂���B " & WScript.Arguments.Unnamed(0)
          Exit Function
       End If

    Else
       '�����Ώۃt�H���_���ȗ������Ƃ��́A�J�����g�f�B���N�g��������
       strExecFolder = getCurrentDir()
    End If

    '�X�C�b�`�����擾
    lngSwitchCount = WScript.Arguments.Named.Count

    '/t �X�C�b�`�i�w�b�_�[�s�j�̃`�F�b�N
    If WScript.Arguments.Named.Exists("t") = False Then
       blnHeader = False
    Else
       blnHeader = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/f �X�C�b�`�i�t�@�C�����̂ݕ\���j�̃`�F�b�N
    If WScript.Arguments.Named.Exists("f") = False Then
       blnFileName = False
    Else
       blnFileName = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/s �X�C�b�`�i�t�@�C���T�C�Y�\���j�̃`�F�b�N
    If WScript.Arguments.Named.Exists("s") = False Then
       blnFileSize = False
    Else
       blnFileSize = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/c /u /a(�쐬�E�X�V�E�A�N�Z�X��)�X�C�b�`�̃`�F�b�N
    ':�쐬���őI�ʃX�C�b�`
    If WScript.Arguments.Named.Exists("c") = False Then
       blnFileCreate = False
    Else
       blnFileCreate = True
       lngSwitchCount = lngSwitchCount - 1
    End If
    ':�X�V���őI�ʃX�C�b�`
    If WScript.Arguments.Named.Exists("u") = False Then
       blnFileUpdate = False
    Else
       blnFileUpdate = True
       lngSwitchCount = lngSwitchCount - 1
    End If
    ':�ŏI�A�N�Z�X���őI�ʃX�C�b�`
    If WScript.Arguments.Named.Exists("a") = False Then
       blnFileAccess = False
    Else
       blnFileAccess = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/zip (���Ƀt�@�C�����̌���)�X�C�b�`�̃`�F�b�N
    If WScript.Arguments.Named.Exists("zip") = False Then
       blnZipFile = False
    Else
       blnZipFile = True
       lngSwitchCount = lngSwitchCount - 1
    End If

    '/sub (�T�u�t�H���_����)�X�C�b�`�̃`�F�b�N
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

    '�t�@�C���ꗗ���擾���܂�
    If blnSubFolder = False Then
       strList = SearchFolder(strExecFolder)     '/sub���w�莞
    Else
       strList = sSearchFolderAll(strExecFolder) '/sub�w�莞
    End If

    '�w�b�_�[�s�̏o��
    If blnHeader = True Then

       strPrintData = "�t�@�C����"

       '�t�@�C���T�C�Y
       If blnFileSize = True Then
          strPrintData = strPrintData _
                       & cDelimiter _
                       & "�T�C�Y"
       End If
       '�t�@�C���쐬��
       If blnFileCreate = True Then
          strPrintData = strPrintData _
                       & cDelimiter _
                       & "�쐬��"
       End If
       '�t�@�C���X�V��
       If blnFileUpdate = True Then
          strPrintData = strPrintData _
                       & cDelimiter _
                       & "�X�V��"
       End If
       '�t�@�C���ŏI�A�N�Z�X��
       If blnFileAccess = True Then
          strPrintData = strPrintData _
                       & cDelimiter _
                       & "�ŏI�A�N�Z�X��"
       End If

       WScript.Echo strPrintData

    End If

    '[POINT!]�t�@�C�����X�g�̏���
    For Each varFilename In strList

        strPrintData = vbNullString

        '�p�X���\�����邩�t�@�C�����݂̂�\�����邩
        If blnFileName = True Then
           strPrintData = strPrintData _
                        & objFS.GetFilename(varFilename)
        Else
           strPrintData = strPrintData _
                            & varFilename
        End If

        '�t�@�C���T�C�Y���擾
        '[POINT!]�擾�ł��Ȃ��Ƃ��ivarFileSize��vbNullString�̂Ƃ�)�͏��Ƀt�@�C�����̃t�@�C��
        On Error Resume Next
        varFileSize = vbNullString
        varFileSize = objFS.GetFile(varFilename).Size
        On Error GoTo 0

        If blnFileSize = True Then
           strPrintData = strPrintData _
                        & cDelimiter _
                        & varFileSize
        End If

        '�t�@�C���쐬��
        If blnFileCreate = True Then
           On Error Resume Next
           dteDateCreated = vbNullString
           dteDateCreated = objFS.GetFile(varFilename).DateCreated
           On Error GoTo 0
           strPrintData = strPrintData _
                        & cDelimiter _
                        & dteDateCreated
        End If

        '�t�@�C���X�V��
        If blnFileUpdate = True Then
           On Error Resume Next
           dteDateLastModified = vbNullString
           dteDateLastModified = objFS.GetFile(varFilename).DateLastModified
           On Error GoTo 0
           strPrintData = strPrintData _
                        & cDelimiter _
                        & dteDateLastModified
        End If

        '�t�@�C���ŏI�A�N�Z�X��
        If blnFileAccess = True Then
           On Error Resume Next
           dteDateLastAccessed = vbNullString
           dteDateLastAccessed = objFS.GetFile(varFilename).DateLastAccessed
           On Error GoTo 0
           strPrintData = strPrintData _
                        & cDelimiter _
                        & dteDateLastAccessed
        End If

        '�t�@�C������\��
        ':�t�@�C���T�C�Y���擾�ł�����(zip���̃t�@�C���ȊO)�̂ݕ\��
        ':�������A/zip�X�C�b�`������Ƃ��́A�������ɕ\��
        If varFileSize <> vbNullString Or _
          (blnZipFile  = True         ) Then

           WScript.Echo strPrintData

        End If

    Next

    '�I�u�W�F�N�g�̔j��
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    FolderSearch = True

End Function
