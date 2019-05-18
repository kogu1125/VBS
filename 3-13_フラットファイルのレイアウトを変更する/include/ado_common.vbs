Option Explicit

'==========================================================
'= テキストファイル用のADOコネクションを作成
'==========================================================
Function getTextConnection(ByRef tmpCon, ByVal tmpPath)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'テキストファイル用のADOコネクションを作成します
    tmpCon.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
                "DBQ=" & tmpPath & ";" & _
                "ReadOnly=0"

    'コネクションの結果を判定
    If Err.Number = 0 Then
       getTextConnection = True  'コネクション成功
    Else
       getTextConnection = False 'コネクションエラー（フォルダ無し）
    End if

End Function

'==========================================================
'= Excelワークブック用のADOコネクションを作成
'==========================================================
Function getExcelConnection(ByRef tmpCon, ByVal tmpPath)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'Excelワークブック用のADOコネクションを作成します
    tmpCon.Open "Driver={Microsoft Excel Driver (*.xls)};" & _
                "DBQ=" & tmpPath & ";" & _
                "ReadOnly=0"

    'コネクションの結果を判定
    If Err.Number = 0 Then
       getExcelConnection = True  'コネクション成功
    Else
       getExcelConnection = False 'コネクションエラー（ファイル無し）
    End if

End Function

'==========================================================
'= Accessデータベース(MDB)用のADOコネクションを作成
'==========================================================
Function getMDBConnection(ByRef tmpCon, ByVal tmpPath)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'MDB用のADOコネクションを作成します
    tmpCon.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
                "DBQ=" & tmpPath & ";"

    'コネクションの結果を判定
    If Err.Number = 0 Then
       getMDBConnection = True  'コネクション成功
    Else
       getMDBConnection = False 'コネクションエラー（MDBファイル無し）
    End if

End Function

'==========================================================
'= SQL Server用のADOコネクションを作成
'==========================================================
Function getSQLSvConnection(ByRef tmpCon, ByVal tmpSVName, ByVal tmpDBName, ByVal tmpUID, ByVal tmpPWD)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'SQL Server用のADOコネクションを作成します
    tmpCon.Open "Driver={SQL Server};" _
              & "server=" & tmpSVName & ";" _
              & "database=" & tmpDBName & ";" _
              & "uid=" & tmpUID & "; pwd=" & tmpPWD & ";"

    'コネクションの結果を判定
    If Err.Number = 0 Then
       getSQLSvConnection = True  'コネクション成功
    Else
       getSQLSvConnection = False 'コネクションエラー
    End if

End Function

'==========================================================
'= MySQL用のADOコネクションを作成
'==========================================================
Function getMySQLConnection(ByRef tmpCon, ByVal tmpSVName, ByVal tmpDBName, ByVal tmpUID, ByVal tmpPWD)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'MySQL用のADOコネクションを作成します
    tmpCon.Open "Driver={MySQL ODBC 3.51 Driver};" _
              & "server=" & tmpSVName & ";" _
              & "database=" & tmpDBName & ";" _
              & "uid=" & tmpUID & "; pwd=" & tmpPWD & ";"

    'コネクションの結果を判定
    If Err.Number = 0 Then
       getMySQLConnection = True  'コネクション成功
    Else
       getMySQLConnection = False 'コネクションエラー
    End if

End Function

'==========================================================
'= PostgreSQL用のADOコネクションを作成
'==========================================================
Function getPSQLConnection(ByRef tmpCon, ByVal tmpSVName, ByVal tmpDBName, ByVal tmpUID, ByVal tmpPWD)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'PostgreSQL用のADOコネクションを作成します
    tmpCon.Open "Driver={PostgreSQL};" _
              & "server=" & tmpSVName & ";" _
              & "database=" & tmpDBName & ";" _
              & "username=" & tmpUID & "; password=" & tmpPWD & ";"

    'コネクションの結果を判定
    If Err.Number = 0 Then
       getPSQLConnection = True  'コネクション成功
    Else
       getPSQLConnection = False 'コネクションエラー
    End if

End Function

'==========================================================
'= ORACLE用のADOコネクションを作成
'==========================================================
Function getORACLEConnection(ByRef tmpCon, ByVal tmpDBName, ByVal tmpUID, ByVal tmpPWD)

    Set tmpCon = Createobject("ADODB.Connection")

    On Error Resume Next

    'ORACLE用のADOコネクションを作成します
    tmpCon.Open "Driver={ORACLE ODBC DRIVER};" _
              & "DBQ=" & tmpDBName & ";" _
              & "uid=" & tmpUID & "; pwd=" & tmpPWD & ";"

    'コネクションの結果を判定
    If Err.Number = 0 Then
       getORACLEConnection = True  'コネクション成功
    Else
       getORACLEConnection = False 'コネクションエラー
    End if

End Function
