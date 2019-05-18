Option Explicit

Const cWait = 1000 '�`�F�b�N�Ԋu(100�`1000�𐄏�)

Function WaitFileUpdate()

    Dim objFS
    Dim strTargetFile
    Dim lngWait
    Dim lngSize
    Dim lngSizeBefore
    Dim dteStartTime
    Dim lngLapTime
    Dim blnDispTime
    Dim lngSwitchCount

    '�v���V�[�W���̌��ʂ����������܂�
    WaitFileUpdate = False

    '�R�}���h���C�������őΏۂ̃t�@�C�����w�肳��Ă��邩��
    '�`�F�b�N���܂�
    If WScript.Arguments.Unnamed.Count = 2 Then

       '�^�[�Q�b�g�t�@�C�����ۑ������t�H���_�̗L�����`�F�b�N���܂�
       '�������A�˗��p�X���ɂ̓t�@�C�������܂݂܂�
       If FileCheck(WScript.Arguments(0), strTargetFile) = False Then
          '�p�X��������Ȃ������Ƃ��̓G���[���b�Z�[�W��\�����v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �^�[�Q�b�g�t�@�C�������݂��܂���B " & WScript.Arguments.Unnamed(0)
          Exit Function
       End If

       '�`�F�b�N�b���p�����^���擾
       lngWait = WScript.Arguments.Unnamed(1)

       '�`�F�b�N�b�������l���ǂ������`�F�b�N���܂�
       If IsNumeric(lngWait) = True Then

          lngWait = CLng(lngWait) '�O�̂��ߐ����^�ɕϊ����܂�

          '�`�F�b�N�b����1�����̂Ƃ��̓p�����^�G���[�Ńv���V�[�W�����I��
          If lngWait < 1 Then
             WScript.Echo "ERROR : �b���ɂ�0�ȏ���w�肵�Ă��������B"
             Exit Function
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

    '�^�[�Q�b�g�t�@�C���̃T�C�Y�i�����l�j
    lngSizeBefore = objFS.GetFile(strTargetFile).Size

    WScript.echo "�`�F�b�N���J�n���܂��B " & strTargetFile

    dteStartTime = Now()

    '�ҋ@����
    Do

       '�t�@�C���T�C�Y���擾
       lngSize = objFS.GetFile(strTargetFile).Size

       '�X�V����Ă��邩�ǂ������t�@�C���T�C�Y����`�F�b�N���܂�
       If lngSize <> lngSizeBefore Then
          '�X�V����Ă���Ƃ��̓`�F�b�N�J�n���������݂ɂ���
          dteStartTime = Now()
       End If

       '�^�[�Q�b�g�t�@�C���̃T�C�Y����ۑ�
       lngSizeBefore = lngSize

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

       WScript.Sleep(cWait) '1�b�ҋ@

    Loop

    WScript.echo lngWait & "�b�ԍX�V������܂���ł����B�`�F�b�N�I�����܂��B"

    '�I�u�W�F�N�g�̔j��
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    WaitFileUpdate = True

End Function

'�c�莞�Ԃ̕\��
Sub DispTimeLeft(ByVal pTime)

    WScript.StdOut.Write "�c�� " & pTime & "�b    " & vbCr

End Sub
