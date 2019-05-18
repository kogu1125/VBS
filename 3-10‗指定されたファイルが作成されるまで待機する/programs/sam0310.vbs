Option Explicit

Const cWait = 1000 '�`�F�b�N�Ԋu(100�`1000�𐄏�)

Function WaitFileExists()

    Dim objFS
    Dim strTargetFile
    Dim lngWait
    Dim dteStartTime
    Dim blnExists
    Dim lngLapTime
    Dim blnDispTime
    Dim lngSwitchCount

    '�v���V�[�W���̌��ʂ����������܂�
    WaitFileExists = False

    '�R�}���h���C�������őΏۂ̃t�@�C�����w�肳��Ă��邩��
    '�`�F�b�N���܂�
    If WScript.Arguments.Unnamed.Count = 2 Then

       '�^�[�Q�b�g�t�@�C�����ۑ������t�H���_�̗L�����`�F�b�N���܂�
       '�������A�˗��p�X���ɂ̓t�@�C�������܂݂܂�
       If FileFolderCheck(WScript.Arguments.Unnamed(0), strTargetFile) = False Then
          '�p�X��������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �^�[�Q�b�g�̃p�X�����݂��܂���B " & WScript.Arguments.Unnamed(0)
          Exit Function
       End If

       '�ҋ@�b���p�����^���擾
       lngWait = WScript.Arguments.Unnamed(1)

       '�b�������l���ǂ������`�F�b�N���܂�
       If IsNumeric(lngWait) = True Then

          lngWait = CLng(lngWait) '�O�̂��ߐ����^�ɕϊ����܂�

          '�b����1�����̂Ƃ��̓p�����^�G���[�Ńv���V�[�W�����I��
          If lngWait < 1 Then
             WScript.Echo "ERROR : �b���ɂ�1�ȏ���w�肵�Ă��������B"
             Exit Function
          End If
       Else
          WScript.Echo "ERROR : �b���͐��l�Ŏw�肵�Ă��������B"
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

    '/d �X�C�b�`�i�c��b���̕\���j
    If WScript.Arguments.Named.Exists("d") = True Then
       blnDispTime = True
       lngSwitchCount = lngSwitchCount - 1
    Else
       blnDispTime = False
    End If

    '�X�C�b�`��ނ̃`�F�b�N
    If lngSwitchCount > 0 Then '�]���ȃX�C�b�`������
       '�v���V�[�W�����I�����܂�
       WScript.Echo "ERROR : �����ȃX�C�b�`������܂��B"
       Exit Function
    End If

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    blnExists = False

    '�����ɑҋ@����Ƃ��͎c��b����\�����Ȃ��悤�ɂ���
    If lngWait = 0 Then
       blnDispTime = False
    End If

    '�J�n����
    dteStartTime = Now()

    '�ҋ@����
    Do

       '�^�[�Q�b�g�t�@�C���̗L�����`�F�b�N���A���������Ƃ���
       '�ҋ@�������I�����܂�
       If objFS.FileExists(strTargetFile) = True Then
          blnExists = True
          Exit Do
       End If

       lngLapTime = DateDiff("s", dteStartTime, Now())

       '/d�X�C�b�`������Ƃ��́A�c��b����\�����܂�
       If blnDispTime = True Then
          DispTimeLeft(lngWait - lngLapTime)
       End If

       '���Ԑ؂�̔���i�ҋ@�����J�n����̌o�ߕb���Ōv��j
       If lngLapTime >= lngWait And _
          lngWait    > 0        Then

          Exit Do

       End If

       WScript.Sleep(cWait) '�ҋ@

    Loop

    '�I�u�W�F�N�g�̔j��
    Set objFS = Nothing

    '���ʂ̕\��
    If blnExists = False Then
       WScript.Echo "�쐬����܂���ł��� " & strTargetFile
       Exit Function
    Else
       WScript.Echo "�쐬����܂��� " & strTargetFile
    End If

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    WaitFileExists = True

End Function

'�c�莞�Ԃ̕\��
Sub DispTimeLeft(ByVal pTime)

    WScript.StdOut.Write "�c�� " & pTime & "�b    " & vbCr

End Sub
