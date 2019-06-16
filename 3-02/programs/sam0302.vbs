Option Explicit

Function TimestampCopy()

    Dim objWSArg
    Dim objFS
    Dim strCopyFrom
    Dim strCopyTo
    Dim strFileName
    Dim strPrefix

    '�v���V�[�W���̌��ʂ����������܂�
    TimestampCopy = False

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    'objWSArg�I�u�W�F�N�g�̃C���X�^���X���R�s�[
    Set objWSArg = WScript.Arguments

    '�R�}���h���C�������ŃR�s�[���t�@�C�����ƃR�s�[��t�H���_����
    '�w�肳��Ă��邩���`�F�b�N���܂�
    If objWSArg.Unnamed.Count = 2 Then

       '�R�s�[���t�@�C���̗L�����`�F�b�N���܂�
       If FileCheck(objWSArg.Unnamed(0), strCopyFrom) = False Then
          WScript.Echo "�R�s�[����t�@�C�� " & objWSArg.Unnamed(0) & " �́A���݂��܂���B"
          Exit Function
       End If

       '�R�s�[��t�H���_�̗L�����`�F�b�N���܂�
       If FolderCheck(objWSArg.Unnamed(1), strCopyTo) = False Then
          '����ł�������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "�R�s�[��t�H���_ " & objWSArg.Unnamed(1) & " �́A���݂��܂���B"
          Exit Function
       End If

    Else
       '�p�����^���������w�肳��Ă��Ȃ��Ƃ��́A�G���[���b�Z�[�W��\����
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �p�����^�𐳂����w�肵�Ă��������B"
       Exit Function
    End If

    '[POINT!]�ǉ�������t�����̕ҏW
    If objWSArg.Named.Count = 0 Then ' �X�C�b�`���ȗ������Ƃ���yyyymmdd
       strPrefix = Right("000" & Year(Date), 4) & _
                   Right("00" & Month(Date), 2) & _
                   Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D1") = True Then ' yyyymmdd
           strPrefix = Right("000" & Year(Date), 4) & _
                       Right("00" & Month(Date), 2) & _
                       Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D2") = True Then ' yyyymm
           strPrefix = Right("000" & Year(Date), 4) & _
                       Right("00" & Month(Date), 2)

    ElseIf objWSArg.Named.Exists("D3") = True Then ' mmdd
           strPrefix = Right("00" & Month(Date), 2) & _
                       Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("D4") = True Then ' yyyy
           strPrefix = Right("000" & Year(Date), 4)

    ElseIf objWSArg.Named.Exists("D5") = True Then ' mm
           strPrefix = Right("00" & Month(Date), 2)

    ElseIf objWSArg.Named.Exists("D6") = True Then ' dd
           strPrefix = Right("00" & Day(Date), 2)

    ElseIf objWSArg.Named.Exists("T1") = True Then 'hhmmss
           strPrefix = Right("00" & Hour(Time), 2) & _
                       Right("00" & Minute(Time), 2) & _
                       Right("00" & Second(Time), 2)

    ElseIf objWSArg.Named.Exists("T2") = True Then ' hhmm
           strPrefix = Right("00" & Hour(Time), 2) & _
                       Right("00" & Minute(Time), 2)

    ElseIf objWSArg.Named.Exists("T3") = True Then ' mmss
           strPrefix = Right("00" & Minute(Time), 2) & _
                       Right("00" & Second(Time), 2)

    ElseIf objWSArg.Named.Exists("T4") = True Then ' yyyymm
           strPrefix = Right("00" & Hour(Time), 2)

    ElseIf objWSArg.Named.Exists("T5") = True Then ' mm
           strPrefix = Right("00" & Minute(Time), 2)

    ElseIf objWSArg.Named.Exists("T6") = True Then ' ss
           strPrefix = Right("00" & Second(Time), 2)

    Else
       '�s���ȏ����X�C�b�`�̓G���[
       WScript.Echo "ERROR : �����𐳂����w�肵�Ă��������B"
       Exit Function
    End If

    '[POINT!]�R�s�[��̃t�@�C�����ƃp�X��ҏW���܂�(yyyymmdd_�t�@�C����)
    strFileName = strPrefix & "_" & objFS.GetFilename(strCopyFrom)
    strCopyTo = objFS.BuildPath(strCopyTo, strFileName)

    '[POINT!]�t�@�C�����R�s�[���܂�
    '�������A���ɓ������O�̃t�@�C��������Ƃ��͍쐬���܂���
    If objFS.FileExists(strCopyTo) = False Then
       objFS.CopyFile strCopyFrom, strCopyTo
       WScript.Echo "�t�@�C�����R�s�[���܂���"
       WScript.Echo strCopyTo
    Else
       WScript.Echo "�������O�̃t�@�C�������݂��܂�"
       WScript.Echo strCopyTo
    End If

    '�I�u�W�F�N�g�̔j��
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    TimestampCopy = True

End Function
