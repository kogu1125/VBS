Option Explicit

Function FileRename()

    Dim objWSArg
    Dim objFS
    Dim strExecFolder
    Dim strFilePath
    Dim strFilename
    Dim strFileBasename
    Dim strFileExt
    Dim strFilenameAfter
    Dim strNewFoldername
    Dim strNewFilename
    Dim strNewFile
    Dim strList
    Dim strListItem
    Dim strBefore
    Dim strAfter
    Dim blnNull
    Dim lngCount
    Dim blnSubFolder
    Dim lngSwitchCount

    '�v���V�[�W���̌��ʂ����������܂�
    FileRename = False

    'objWSArg�I�u�W�F�N�g�̃C���X�^���X���R�s�[
    Set objWSArg = WScript.Arguments

    lngSwitchCount = objWSArg.Named.Count

    '�u���������null�ɂ���ꍇ�̃X�C�b�`
    If objWSArg.Named.Exists("null") = False Then
       blnNull = False
    Else
       '�J�b�g�̎w�莞�͒u���������vbNullString��
       strAfter = vbNullString
       blnNull  = True

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

    '�R�}���h���C�������Ń^�[�Q�b�g�t�H���_�ƒu���������e��
    '�w�肳��Ă��邩���`�F�b�N���܂�
    If objWSArg.Unnamed.Count = 3 Or _
       (blnNull = True And objWSArg.Unnamed.Count = 2)   Then

       '�����t�H���_�̗L�����`�F�b�N���܂�
       If FolderCheck(WScript.Arguments(0), strExecFolder) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\����
          '�v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �w�肵���t�H���_�͑��݂��܂���B " & objWSArg.Unnamed(0)
          Exit Function
       End If

       strBefore =objWSArg.Unnamed(1) '�u�������O

       '�u���������/null�X�C�b�`�������ꍇ�̂ݗL��
       If blnNull = False Then
          strAfter = objWSArg.Unnamed(2) '�u��������
       End If

       '�u�������O�オ�����ꍇ�͏����ł��Ȃ�
       If UCase(strBefore) = UCase(strAfter) Then
          WScript.Echo "ERROR : �u�������O��̓��e�������ł��B "
          Exit Function
       End If

    Else
       '�p�����^���������w�肳��Ă��Ȃ��Ƃ��́A�G���[���b�Z�[�W��\����
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �p�����^�𐳂����w�肵�Ă��������B"
       Exit Function
    End If

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '�t�@�C�����X�g�̎擾
    If blnSubFolder = False Then
       strList = SearchFolder(strExecFolder)     '/sub���w�莞
    Else
       strList = sSearchFolderAll(strExecFolder) '/sub�w�莞
    End If

    '�ϐ��̏�����
    lngCount = 0

    '[POINT!]�t�@�C�����X�g�̏���
    For Each strListItem In strList

        '���o�����A�C�e������A�t�H���_�̃p�X���擾
        strNewFoldername = getFilePath(strListItem)
        '���o�����A�C�e������A�t�@�C�����݂̂��擾
        strFilename      = objFS.GetFilename(strListItem)
        '�g���q���܂܂Ȃ��t�@�C�����i�T���v�������p�j
        strFileBasename  = objFS.GetBasename(strListItem)
        '�g���q�i�T���v�������p�j
        strFileExt       = objFS.GetExtensionName(strListItem)

        '���l�[���̑Ώۃt�@�C�����ǂ����𔻒f���܂�
        '�Ώۃt�@�C���ł���΃t�@�C���������l�[�����܂�
        If InStr(1, strFilename, strBefore, vbTextCompare) > 0 Then

           lngCount = lngCount + 1 '�T���v�������p�i�ʔԃJ�E���g�j

           '���l�[����̃t�@�C�����ƃp�X��ҏW���܂�
           strNewFilename = Replace(strFilename, strBefore, strAfter, 1, -1, vbTextCompare)
           strNewFile = objFS.BuildPath(strNewFoldername, strNewFilename)

           '�������O�̃t�@�C��������Ƃ��͏����𒆒i����
           If objFS.FileExists(strNewFile) = True Then
              WScript.echo "�����̃t�@�C�������݂��܂� -> " & strNewFilename
           Else
              'FileMove���\�b�h���g���ăt�@�C������ύX���܂�
              objFS.MoveFile strListItem, strNewFile

              WScript.echo "�t�@�C������ύX���܂��� �O: " & strFilename & _
                           " ��: " & strNewFilename

           End If

        End If

    Next

    WScript.echo "�������������܂���"

    '�I�u�W�F�N�g�̔j��
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    FileRename = True

End Function
