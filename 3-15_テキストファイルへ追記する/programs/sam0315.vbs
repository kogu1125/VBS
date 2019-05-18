Option Explicit

Function LogOutput()

    Dim objWSArg
    Dim objFS
    Dim strExecFolder
    Dim strFilePath
    Dim strFilename
    Dim strFilenameAfter
    Dim strNewFilename
    Dim strNewFile
    Dim strList
    Dim strListItem
    Dim strBefore
    Dim strAfter
    Dim blnNull
    Dim lngCount
    Dim strLogFile
    Dim strMsg

    lngCount = 0

    '�v���V�[�W���̌��ʂ����������܂�
    LogOutput = False

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    '���O�t�@�C���̏ꏊ
    strLogFile = objFS.BuildPath(getCurrentDir(), "sam0315.log")

    'objWSArg�I�u�W�F�N�g�̃C���X�^���X���R�s�[
    Set objWSArg = WScript.Arguments

    '�u�������オnull�̏ꍇ�̓X�C�b�`�ōs��
    If objWSArg.Named.Count = 1 Then

       '�X�C�b�`���������w�肳��Ă���ꍇ
       If objWSArg.Named.Exists("null") = True Then

          '�J�b�g�̎w�莞�͒u���������vbNullString��
          strAfter = vbNullString

       '�X�C�b�`���ԈႦ�Ďw�肵���ꍇ
       Else
          strMsg = "ERROR : /null �ȊO�̃X�C�b�`���w�肳��Ă��܂��B"
          Call TextWriteBottom(strLogFile, strMsg)
          Exit Function
       End If

       blnNull = True

    Else
       blnNull = False
    End If

    '�R�}���h���C�������Ń^�[�Q�b�g�t�H���_�ƒu���������e��
    '�w�肳��Ă��邩���`�F�b�N���܂�
    If objWSArg.Unnamed.Count = 3 Or _
       (blnNull = True And objWSArg.Unnamed.Count = 2)   Then

       '�����t�H���_�̗L�����`�F�b�N���܂�
       If FolderCheck(WScript.Arguments(0), strExecFolder) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\����
          '�v���V�[�W�����I�����܂�
          strMsg = "ERROR : �w�肵���t�H���_�͑��݂��܂���B " & objWSArg.Unnamed(0)
          Call TextWriteBottom(strLogFile, strMsg)
          Exit Function
       End If

       strBefore =objWSArg.Unnamed(1) '�u�������O

       '�u���������/null�X�C�b�`�������ꍇ�̂ݗL��
       If blnNull = False Then
          strAfter = objWSArg.Unnamed(2) '�u��������
       End If

       '�u�������O�オ�����ꍇ�͏����ł��Ȃ�
       If UCase(strBefore) = UCase(strAfter) Then
          strMsg = "ERROR : �u�������O��̓��e�������ł��B "
          Call TextWriteBottom(strLogFile, strMsg)
          Exit Function
       End If

    Else
       '�p�����^���������w�肳��Ă��Ȃ��Ƃ��́A�G���[���b�Z�[�W��\����
       '�v���V�[�W�����I�����܂�
       strMsg = "ERROR : �p�����^�𐳂����w�肵�Ă��������B"
       Call TextWriteBottom(strLogFile, strMsg)
       Exit Function
    End If

    '�t�@�C�����X�g�̎擾
    strList = SearchFolder(strExecFolder)

    '[POINT!]�t�@�C�����X�g�̏���
    For Each strListItem In strList

        '���o�����A�C�e������A�t�@�C�����݂̂��擾
        strFilename = objFS.GetFilename(strListItem)

        '���l�[���̑Ώۃt�@�C�����ǂ����𔻒f���܂�
        '�Ώۃt�@�C���ł���΃t�@�C���������l�[�����܂�
        If InStr(1, strFilename, strBefore, vbTextCompare) > 0 Then

           lngCount = lngCount + 1 '�T���v�������p�i�ʔԃJ�E���g�j

           '���l�[����̃t�@�C�����ƃp�X��ҏW���܂�
           strNewFilename = Replace(strFilename, strBefore, strAfter, 1, -1, vbTextCompare)
           strNewFile = objFS.BuildPath(strExecFolder, strNewFilename)

           '�������O�̃t�@�C��������Ƃ��͏����𒆒i����
           If objFS.FileExists(strNewFile) = True Then
              strMsg = "�����̃t�@�C�������݂��܂� -> " & strNewFilename
              Call TextWriteBottom(strLogFile, strMsg)
           Else
              'FileMove���\�b�h���g���ăt�@�C������ύX���܂�
              objFS.MoveFile strListItem, strNewFile

              strMsg = "�t�@�C������ύX���܂��� �O: " & strFilename & _
                       " ��: " & strNewFilename
              Call TextWriteBottom(strLogFile, strMsg)

           End If

        End If

    Next

    strMsg = "�������������܂���"
    Call TextWriteBottom(strLogFile, strMsg)

    '�I�u�W�F�N�g�̔j��
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    LogOutput = True

End Function
