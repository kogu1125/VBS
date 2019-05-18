Option Explicit

Const cWait = 1000 'チェック間隔(100～1000を推奨)

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

    'プロシージャの結果を初期化します
    WaitFileUpdate = False

    'コマンドライン引数で対象のファイルが指定されているかを
    'チェックします
    If WScript.Arguments.Unnamed.Count = 2 Then

       'ターゲットファイルが保存されるフォルダの有無をチェックします
       'ただし、依頼パス情報にはファイル名も含みます
       If FileCheck(WScript.Arguments(0), strTargetFile) = False Then
          'パスが見つからなかったときはエラーメッセージを表示しプロシージャを終了します
          WScript.Echo "ERROR : ターゲットファイルが存在しません。 " & WScript.Arguments.Unnamed(0)
          Exit Function
       End If

       'チェック秒数パラメタを取得
       lngWait = WScript.Arguments.Unnamed(1)

       'チェック秒数が数値かどうかをチェックします
       If IsNumeric(lngWait) = True Then

          lngWait = CLng(lngWait) '念のため整数型に変換します

          'チェック秒数が1未満のときはパラメタエラーでプロシージャを終了
          If lngWait < 1 Then
             WScript.Echo "ERROR : 秒数には0以上を指定してください。"
             Exit Function
          End If

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

    'ターゲットファイルのサイズ（初期値）
    lngSizeBefore = objFS.GetFile(strTargetFile).Size

    WScript.echo "チェックを開始します。 " & strTargetFile

    dteStartTime = Now()

    '待機処理
    Do

       'ファイルサイズを取得
       lngSize = objFS.GetFile(strTargetFile).Size

       '更新されているかどうかをファイルサイズからチェックします
       If lngSize <> lngSizeBefore Then
          '更新されているときはチェック開始時刻を現在にする
          dteStartTime = Now()
       End If

       'ターゲットファイルのサイズ情報を保存
       lngSizeBefore = lngSize

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

       WScript.Sleep(cWait) '1秒待機

    Loop

    WScript.echo lngWait & "秒間更新がありませんでした。チェック終了します。"

    'オブジェクトの破棄
    Set objFS = Nothing

    'プロシージャの結果をTrueにします
    WaitFileUpdate = True

End Function

'残り時間の表示
Sub DispTimeLeft(ByVal pTime)

    WScript.StdOut.Write "残り " & pTime & "秒    " & vbCr

End Sub
