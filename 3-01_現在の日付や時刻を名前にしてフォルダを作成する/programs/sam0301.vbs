Option Explicit

Function CreateFolder()

    Dim objWSArg
    Dim objFS
    Dim strFolderName
    Dim strCreatePath

    '�v���V�[�W���̌��ʂ����������܂�
    CreateFolder = False

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'objWSArg�I�u�W�F�N�g�̃C���X�^���X���R�s�[
    Set objWSArg = WScript.Arguments

    '�R�}���h���C�������ō쐬�悪�w�肳��Ă��邩���`�F�b�N���܂�
    If objWSArg.Unnamed.Count > 0 Then

       '�p�X�̃`�F�b�N�B������Ȃ��Ƃ��̓v���V�[�W�����I��
       If FolderCheck(objWSArg.Unnamed(0), strCreatePath) = False Then
          WScript.Echo "ERROR : �p�X�����݂��܂���B " & objWSArg.Unnamed(0)
          Exit Function
       End If

    Else
       '�R�}���h���C���������w�肳��Ă��Ȃ��Ƃ��̓J�����g�f�B���N�g�����W���p�X
       strCreatePath = getCurrentDir()
    End If

    '[POINT!]�p�����^�̓��e�ɂ���ē��t�⎞�����o�͂��܂�
    If objWSArg.Named.Count = 0 Then ' �X�C�b�`���ȗ������Ƃ���yyyymmdd
       strFolderName = Right("000" & Year(Date), 4) & _
                       Right("00" & Month(Date), 2) & _
                       Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D1") = True Then ' yyyymmdd
           strFolderName = Right("000" & Year(Date), 4) & _
                           Right("00" & Month(Date), 2) & _
                           Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D2") = True Then ' yyyymm
           strFolderName = Right("000" & Year(Date), 4) & _
                           Right("00" & Month(Date), 2)

    ElseIf objWSArg.Named.Exists("D3") = True Then ' mmdd
           strFolderName = Right("00" & Month(Date), 2) & _
                           Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D4") = True Then ' yyyy
           strFolderName = Right("000" & Year(Date), 4)

    ElseIf objWSArg.Named.Exists("D5") = True Then ' mm
           strFolderName = Right("00" & Month(Date), 2)

    ElseIf objWSArg.Named.Exists("D6") = True Then ' dd
           strFolderName = Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("T1") = True Then 'hhmmss
           strFolderName = Right("00" & Hour(Time), 2) & _
                           Right("00" & Minute(Time), 2) & _
                           Right("00" & Second(Time), 2)

    ElseIf objWSArg.Named.Exists("T2") = True Then ' hhmm
           strFolderName = Right("00" & Hour(Time), 2) & _
                           Right("00" & Minute(Time), 2)

    ElseIf objWSArg.Named.Exists("T3") = True Then ' mmss
           strFolderName = Right("00" & Minute(Time), 2) & _
                           Right("00" & Second(Time), 2)

    ElseIf objWSArg.Named.Exists("T4") = True Then ' yyyymm
           strFolderName = Right("00" & Hour(Time), 2)

    ElseIf objWSArg.Named.Exists("T5") = True Then ' mm
           strFolderName = Right("00" & Minute(Time), 2)

    ElseIf objWSArg.Named.Exists("T6") = True Then ' ss
           strFolderName = Right("00" & Second(Time), 2)

    Else
       '�s���ȏ����X�C�b�`�̓G���[
       WScript.Echo "ERROR : �����𐳂����w�肵�Ă��������B"
       Exit Function
    End If

    strCreatePath = objFS.BuildPath(strCreatePath, strFolderName)

    '[POINT!]�t�H���_���쐬���܂�
    '        �������A���ɓ������O�̃t�H���_������Ƃ��͍쐬���܂���
    If objFS.FolderExists(strCreatePath) = False Then
       objFS.CreateFolder strCreatePath
       WScript.Echo "�t�H���_���쐬���܂���"
       WScript.Echo strCreatePath
    Else
       WScript.Echo "ERROR : �������O�̃t�H���_�����݂��܂�"
       WScript.Echo strCreatePath
    End If

    '�I�u�W�F�N�g�̔j��
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    CreateFolder = True

End Function
