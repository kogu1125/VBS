Option Explicit

'==========================================================
'= �t�H���_�̗L���`�F�b�N
'==========================================================
Function FolderCheck(ByVal tmpPath, ByRef strPath)

    Dim objFS
    Dim strCheckPath

    '�v���V�[�W���̌��ʂ����������܂�
    FolderCheck = False

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '�`�F�b�N����t�H���_�̃p�X�����쐬�i��΃p�X�ɕҏW�j
    strCheckPath = objFS.GetAbsolutePathName(tmpPath)

    '�ҏW�����p�X�Ńt�H���_�̗L�����`�F�b�N
    If objFS.FolderExists(strCheckPath) = True Then
       FolderCheck = True
       strPath = strCheckPath '�p�X����Ԃ�
    End If

    Set objFS = Nothing '�I�u�W�F�N�g�̔j��

End Function

'==========================================================
'= �t�@�C�������܂ރp�X�̏ꍇ�̃t�H���_�̗L���`�F�b�N
'==========================================================
Function FileFolderCheck(ByVal tmpPath, ByRef strPath)

    Dim objFS
    Dim objPath
    Dim strCheckPath
    Dim strTmpPath

    '�v���V�[�W���̌��ʂ����������܂�
    FileFolderCheck = False

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '��΃p�X�ɕҏW
    strTmpPath = objFS.GetAbsolutePathName(tmpPath)

    '����Ƀt�@�C���܂ł̃t���p�X����p�X��񂾂����擾
    strCheckPath = Replace(strTmpPath, objFS.GetFilename(strTmpPath), vbNullString)

    '�ҏW�����p�X�Ńt�H���_�̗L�����`�F�b�N
    If objFS.FolderExists(strCheckPath) = True Then
       FileFolderCheck = True
       strPath = strTmpPath '�p�X����Ԃ�
    End If

    Set objFS = Nothing '�I�u�W�F�N�g�̔j��

End Function

'==========================================================
'= �t�@�C�����X�g�̎擾(�w��t�H���_�̂�)
'==========================================================
Function SearchFolder(tmpExecFolder)

    Dim objApl
    Dim objFolder
    Dim objFolderItems
    Dim objItem
    Dim lngCount
    dim strFileList()

    lngCount = 0

    '[POINT!]Shell�I�u�W�F�N�g���쐬���܂�
    Set objApl = CreateObject("Shell.Application")

    '[POINT!]��������t�H���_�̃I�u�W�F�N�g���쐬���܂�
    Set objFolder = objApl.Namespace(tmpExecFolder)

    '[POINT!]�t�H���_�����������܂�
    For Each objItem In objFolder.Items

        '[POINT!]���o���������t�@�C�����ǂ����𔻒f���܂�
        '[POINT!]zip���Ƀt�@�C���́u�t�H���_�v�ŔF�������ꍇ������܂�
        If objItem.IsFolder = False Then
           Redim Preserve strFileList(lngCount)
           strFileList(lngCount) = objItem.Path
           lngCount = lngCount + 1
        End If

    Next

    '�߂�l�ɂ̓t�@�C���̈ꗗ�Ԃ��܂�
    SearchFolder = strFileList

    '�I�u�W�F�N�g�̔j��
    Set objItem = Nothing
    Set objFolderItems = Nothing
    Set objFolder = Nothing
    Set objApl = Nothing

End Function

'==============================================================
'= �t�@�C�����X�g�̎擾(���̊K�w�܂Ō�������)
'==============================================================
Function sSearchFolderAll(tmpExecFolder)

    Dim objApl
    Dim objFolder
    Dim strFileList()

    'Shell�I�u�W�F�N�g���쐬���܂�
    Set objApl = CreateObject("Shell.Application")

    '��������t�H���_�̃I�u�W�F�N�g���쐬���܂�
    Set objFolder = objApl.Namespace(tmpExecFolder)

    '�t�H���_�����������Ăяo���܂�
    Call sSearchFolderAll_Sub(objFolder.Items, strFileList)

    '�߂�l�ɂ̓t�@�C���̈ꗗ�Ԃ��܂�
    sSearchFolderAll = strFileList

    '�I�u�W�F�N�g�̔j��
    Set objFolder = Nothing
    Set objApl = Nothing

End Function

'==============================================================
'= �t�H���_���Ɋ܂܂��t�@�C����t�H���_����������(�ċA�Ăяo��)
'= :sSearchFolderAll�̃T�u���[�`��
'==============================================================
Sub sSearchFolderAll_Sub(ByVal tmpFolderItems, ByRef tmpFileList)

    Dim objFolderItems
    Dim objItem
    Dim lngCount

    '�z��̑傫�����ēx���߂�
    lngCount = 0
    On Error Resume Next
    lngCount = UBound(tmpFileList) + 1
    On Error Goto 0

    '�t�H���_��������
    For Each objItem In tmpFolderItems

        '���o���������t�@�C�����t�H���_���𔻒�
        If objItem.IsFolder Then
           '�t�H���_�ł���΁AItems�I�u�W�F�N�g�����A
           '����������Ƃ���sSearchFolderAll_Sub���u�ċA�Ăяo���v���܂�
           Set objFolderItems = objItem.GetFolder.Items
           Call sSearchFolderAll_Sub(objFolderItems, tmpFileList)

           '�z��̑傫�����ēx���߂�
           lngCount = 0
           On Error Resume Next
           lngCount = UBound(tmpFileList) + 1
           On Error Goto 0

        Else
           '�t�@�C���ł���΁A���X�g�Ɋi�[���܂�
           If Mid(objItem.Path, 2,1) = ":"  Or _
              Mid(objItem.Path, 2,2) = "\\" Then
              ReDim Preserve tmpFileList(lngCount)
              tmpFileList(lngCount) = objItem.Path
              lngCount = lngCount + 1
           End If

        End If

    Next

    '�I�u�W�F�N�g�̔j��
    Set objItem = Nothing
    Set objFolderItems = Nothing

End Sub

'==========================================================
'= �J����zip�t�H���_���쐬
'==========================================================
Function CreateZipMaster(ByVal pDir, ByVal pZipFilename)

    Dim objStream
    Dim objFS

    CreateZipMaster = False

    Set objFS = CreateObject("Scripting.FilesystemObject")
    Set objStream = Createobject("ADODB.Stream")

    '������zip�������ꍇ�̂ݍ쐬
    If objFS.FileExists(objFS.BuildPath(pDir, pZipFilename)) = False Then

       'Zip�쐬
       With objStream
           .Open
           .Type = adTypeText
           .charset = "iso-8859-1"
           .WriteText ChrW(&h50) & ChrW(&h4B) & ChrW(&h05) & ChrW(&h06) _
                    & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) _
                    & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) _
                    & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) & ChrW(&h00) _
                    & ChrW(&h00) & ChrW(&h00) & ChrW(&h00)
           .SaveToFile objFS.BuildPath(pDir, pZipFilename) ,2
           .Close
       End With

       CreateZipMaster = True

    End If

    Set objStream = Nothing
    Set objFS = Nothing

End Function
