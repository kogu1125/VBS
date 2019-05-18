Option Explicit

'==========================================================
'= �e�L�X�g�t�@�C���p��ADO�R�l�N�V�������쐬
'==========================================================
Function getTextConnection(ByRef tmpCon, ByVal tmpPath)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    '�e�L�X�g�t�@�C���p��ADO�R�l�N�V�������쐬���܂�
    tmpCon.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
                "DBQ=" & tmpPath & ";" & _
                "ReadOnly=0"

    '�R�l�N�V�����̌��ʂ𔻒�
    If Err.Number = 0 Then
       getTextConnection = True  '�R�l�N�V��������
    Else
       getTextConnection = False '�R�l�N�V�����G���[�i�t�H���_�����j
    End if

End Function

'==========================================================
'= Excel���[�N�u�b�N�p��ADO�R�l�N�V�������쐬
'==========================================================
Function getExcelConnection(ByRef tmpCon, ByVal tmpPath)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'Excel���[�N�u�b�N�p��ADO�R�l�N�V�������쐬���܂�
    tmpCon.Open "Driver={Microsoft Excel Driver (*.xls)};" & _
                "DBQ=" & tmpPath & ";" & _
                "ReadOnly=0"

    '�R�l�N�V�����̌��ʂ𔻒�
    If Err.Number = 0 Then
       getExcelConnection = True  '�R�l�N�V��������
    Else
       getExcelConnection = False '�R�l�N�V�����G���[�i�t�@�C�������j
    End if

End Function

'==========================================================
'= Access�f�[�^�x�[�X(MDB)�p��ADO�R�l�N�V�������쐬
'==========================================================
Function getMDBConnection(ByRef tmpCon, ByVal tmpPath)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'MDB�p��ADO�R�l�N�V�������쐬���܂�
    tmpCon.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
                "DBQ=" & tmpPath & ";"

    '�R�l�N�V�����̌��ʂ𔻒�
    If Err.Number = 0 Then
       getMDBConnection = True  '�R�l�N�V��������
    Else
       getMDBConnection = False '�R�l�N�V�����G���[�iMDB�t�@�C�������j
    End if

End Function

'==========================================================
'= SQL Server�p��ADO�R�l�N�V�������쐬
'==========================================================
Function getSQLSvConnection(ByRef tmpCon, ByVal tmpSVName, ByVal tmpDBName, ByVal tmpUID, ByVal tmpPWD)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'SQL Server�p��ADO�R�l�N�V�������쐬���܂�
    tmpCon.Open "Driver={SQL Server};" _
              & "server=" & tmpSVName & ";" _
              & "database=" & tmpDBName & ";" _
              & "uid=" & tmpUID & "; pwd=" & tmpPWD & ";"

    '�R�l�N�V�����̌��ʂ𔻒�
    If Err.Number = 0 Then
       getSQLSvConnection = True  '�R�l�N�V��������
    Else
       getSQLSvConnection = False '�R�l�N�V�����G���[
    End if

End Function

'==========================================================
'= MySQL�p��ADO�R�l�N�V�������쐬
'==========================================================
Function getMySQLConnection(ByRef tmpCon, ByVal tmpSVName, ByVal tmpDBName, ByVal tmpUID, ByVal tmpPWD)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'MySQL�p��ADO�R�l�N�V�������쐬���܂�
    tmpCon.Open "Driver={MySQL ODBC 3.51 Driver};" _
              & "server=" & tmpSVName & ";" _
              & "database=" & tmpDBName & ";" _
              & "uid=" & tmpUID & "; pwd=" & tmpPWD & ";"

    '�R�l�N�V�����̌��ʂ𔻒�
    If Err.Number = 0 Then
       getMySQLConnection = True  '�R�l�N�V��������
    Else
       getMySQLConnection = False '�R�l�N�V�����G���[
    End if

End Function

'==========================================================
'= PostgreSQL�p��ADO�R�l�N�V�������쐬
'==========================================================
Function getPSQLConnection(ByRef tmpCon, ByVal tmpSVName, ByVal tmpDBName, ByVal tmpUID, ByVal tmpPWD)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'PostgreSQL�p��ADO�R�l�N�V�������쐬���܂�
    tmpCon.Open "Driver={PostgreSQL};" _
              & "server=" & tmpSVName & ";" _
              & "database=" & tmpDBName & ";" _
              & "username=" & tmpUID & "; password=" & tmpPWD & ";"

    '�R�l�N�V�����̌��ʂ𔻒�
    If Err.Number = 0 Then
       getPSQLConnection = True  '�R�l�N�V��������
    Else
       getPSQLConnection = False '�R�l�N�V�����G���[
    End if

End Function

'==========================================================
'= ORACLE�p��ADO�R�l�N�V�������쐬
'==========================================================
Function getORACLEConnection(ByRef tmpCon, ByVal tmpDBName, ByVal tmpUID, ByVal tmpPWD)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'ORACLE�p��ADO�R�l�N�V�������쐬���܂�
    tmpCon.Open "Driver={ORACLE ODBC DRIVER};" _
              & "DBQ=" & tmpDBName & ";" _
              & "uid=" & tmpUID & "; pwd=" & tmpPWD & ";"

    '�R�l�N�V�����̌��ʂ𔻒�
    If Err.Number = 0 Then
       getORACLEConnection = True  '�R�l�N�V��������
    Else
       getORACLEConnection = False '�R�l�N�V�����G���[
    End if

End Function
