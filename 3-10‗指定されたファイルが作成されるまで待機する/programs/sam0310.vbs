Option Explicit

Const cWait = 1000 'チェック間隔(100～1000を推奨)

Function WaitFileExists()

    Dim objFS
    Dim strTargetFile
    Dim lngWait
    Dim dteStartTime
    Dim blnExists
    Dim lngLapTime
    Dim blnDispTime
    Dim lngSwitchCount

    'プロシージャの結果を初期化します
    WaitFileExists = False

    'コマンドライン引数で対象のファイルが指定されているかを
    'チェックします
    If WScript.Arguments.Unnamed.Count = 2 Then

       'ターゲットファイルが保存されるフォルダの有無をチェックします
       'ただし、依頼パス情報にはファイル名も含みます
       If FileFolderCheck(WScript.Arguments.Unnamed(0), strTargetFile) = False Then
          'パスが見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : ターゲットのパスが存在しません。 " & WScript.Arguments.Unnamed(0)
          Exit Function
       End If

       '待機秒数パラメタを取得
       lngWait = WScript.Arguments.Unnamed(1)

       '秒数が数値かどうかをチェックします
       If IsNumeric(lngWait) = True Then

          lngWait = CLng(lngWait) '念のため整数型に変換します

          '秒数が1未満のときはパラメタエラーでプロシージャを終了
          If lngWait < 1 Then
             WScript.Echo "ERROR : 秒数には1以上を指定してください。"
             Exit Function
          End If
       Else
          WScript.Echo "ERROR : 秒数は数値で指定してください。"
          Exit Function
       End If

    Else
       'パラメタが正しく指定されていないときは、エラーメッセージを表示し
       'プロシージャを終了します
       WScript.Echo "ERROR : パラメタを正しく指定してください。"
       Exit Function
    End If

    'スイッチ数を取得
    lngSwitchCount = WScript.Arguments.Named.Count

    '/d スイッチ（残り秒数の表示）
    If WScript.Arguments.Named.Exists("d") = True Then
       blnDispTime = True
       lngSwitchCount = lngSwitchCount - 1
    Else
       blnDispTime = False
    End If

    'スイッチ種類のチェック
    If lngSwitchCount > 0 Then '余分なスイッチがある
       'プロシージャを終了します
       WScript.Echo "ERROR : 無効なスイッチがあります。"
       Exit Function
    End If

    'ファイルシステムオブジェクトを作成します
    Set objFS = CreateObject("Scripting.FilesystemObject")

    blnExists = False

    '無限に待機するときは残り秒数を表示しないようにする
    If lngWait = 0 Then
       blnDispTime = False
    End If

    '開始時刻
    dteStartTime = Now()

    '待機処理
    Do

       'ターゲットファイルの有無をチェックし、見つかったときは
       '待機処理を終了します
       If objFS.FileExists(strTargetFile) = True Then
          blnExists = True
          Exit Do
       End If

       lngLapTime = DateDiff("s", dteStartTime, Now())

       '/dスイッチがあるときは、残り秒数を表示します
       If blnDispTime = True Then
          DispTimeLeft(lngWait - lngLapTime)
       End If

       '時間切れの判定（待機処理開始からの経過秒数で計る）
       If lngLapTime >= lngWait And _
          lngWait    > 0        Then

          Exit Do

       End If

       WScript.Sleep(cWait) '待機

    Loop

    'オブジェクトの破棄
    Set objFS = Nothing

    '結果の表示
    If blnExists = False Then
       WScript.Echo "作成されませんでした " & strTargetFile
       Exit Function
    Else
       WScript.Echo "作成されました " & strTargetFile
    End If

    'プロシージャの結果をTrueにします
    WaitFileExists = True

End Function

'残り時間の表示
Sub DispTimeLeft(ByVal pTime)

    WScript.StdOut.Write "残り " & pTime & "秒    " & vbCr

End Sub
