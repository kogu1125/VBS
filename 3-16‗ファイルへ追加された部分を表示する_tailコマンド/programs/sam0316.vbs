Option Explicit

Const cWait = 500 '�Ď��Ԋu�Ԋu(100�`1000�𐄏�)

Function TailCommand()

    Dim objFS
    Dim strTargetFile
    Dim lngExec
    Dim objInFile
    Dim strData
    Dim dteStartTime

    '�v���V�[�W���̌��ʂ����������܂�
    TailCommand = False

    '�R�}���h���C�������Ń^�[�Q�b�g�t�@�C����
    '�w�肳��Ă��邩���`�F�b�N���܂�
    If WScript.Arguments.Count = 2 Then

       '�^�[�Q�b�g�t�@�C���̗L�����`�F�b�N���܂�
       If FileCheck(WScript.Arguments(0), strTargetFile) = False Then
          '������Ȃ������Ƃ��̓G���[���b�Z�[�W��\����
          '�v���V�[�W�����I�����܂�
          WScript.Echo "ERROR : �w�肵���t�@�C���͑��݂��܂���B " & WScript.Arguments(0)
          Exit Function
       End If

       '�����b���p�����^���擾
       lngExec = WScript.Arguments(1)

       '�b�������l���ǂ������`�F�b�N���܂�
       If IsNumeric(lngExec) = True Then

          lngExec = CLng(lngExec) '�O�̂��ߐ����^�ɕϊ����܂�

          '�b����0�����̂Ƃ��̓p�����^�G���[�Ńv���V�[�W�����I��
          If lngExec < 0 Then
             WScript.Echo "ERROR : �b���ɂ�0�ȏ���w�肵�Ă��������B"
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

    '�t�@�C���V�X�e���I�u�W�F�N�g���쐬���܂�
    Set objFS = CreateObject("Scripting.FilesystemObject")

    dteStartTime = Now()

    '�Ώۃt�@�C���̏���
    Set objInFile = objFS.OpenTextFile(strTargetFile, ForReading)
    '�J���t�@�C���łȂ��ꍇ�͑S�ǂݍ��݂��s��
    If objInFile.AtEndOfStream = False Then
       objInFile.ReadAll
    End If

    Do

       If objInFile.AtEndOfStream = False Then
          strData = objInFile.ReadAll
          WScript.Stdout.Write strData
       End If

       WScript.Sleep(cWait) '0.5�b�ҋ@

    Loop Until DateDiff("s", dteStartTime, Now()) >= lngExec And _
               lngExec > 0

    objInFile.Close

    '�I�u�W�F�N�g�̔j��
    Set objInFile = Nothing
    Set objFS = Nothing

    '�v���V�[�W���̌��ʂ�True�ɂ��܂�
    TailCommand = True

End Function
