Option Explicit  ' 変数宣言漏れを防ぐ

'-------------------------------------------------------------------------
' 使い方と目的
'-------------------------------------------------------------------------
' 本スクリプトは、Fortnite の設定ファイル (GameUserSettings.ini) を一定間隔で監視し、
' "bDisableMouseAcceleration=False" を見つけた場合に強制的に True に書き換えて
' マウス加速をオフに保つためのものです。
'
' ▼想定配置先
'   %localappdata%\FortniteGame\Saved\Config\WindowsClient\GameUserSettings.ini
'   (自動でパスを取得します)
'
' ▼スタートアップに登録したいとき
'   1) Windowsキー + R で「ファイル名を指定して実行」を開く
'   2) 'shell:startup' と入力しEnter
'   3) 表示されたフォルダに、この .vbs ファイルのショートカットを作成して置く
'      → 次回ログイン時から自動で監視を開始します
'
' ▼終了したいとき
'   - Fortnite側で設定を見つけて通知が出たとき、[いいえ]ボタンで終了できます。
'   - あるいは、タスクマネージャ > wscript.exe を終了させることでも強制終了が可能です。
'
'-------------------------------------------------------------------------
' 変数宣言
'-------------------------------------------------------------------------
Dim fso              ' ファイル操作用 (Scripting.FileSystemObject)
Dim shell            ' シェル操作用 (WScript.Shell)
Dim localAppData     ' %localappdata% のパス展開用
Dim configFilePath   ' Fortnite設定ファイルのパス
Dim SLEEP_INTERVAL   ' 監視間隔(ミリ秒)

SLEEP_INTERVAL = 30000 ' 30秒おきにチェック(1000=1秒, 30000=30秒)

'-------------------------------------------------------------------------
' オブジェクト作成
'-------------------------------------------------------------------------
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

'-------------------------------------------------------------------------
' フォートナイト設定ファイルのパスを組み立てる
'-------------------------------------------------------------------------
localAppData   = shell.ExpandEnvironmentStrings("%localappdata%")  ' 環境変数を展開
configFilePath = localAppData & "\FortniteGame\Saved\Config\WindowsClient\GameUserSettings.ini"

'-------------------------------------------------------------------------
' 起動時の確認メッセージ
'-------------------------------------------------------------------------
Dim startupMsg
startupMsg = "【Fortniteマウス加速 監視スクリプト】" & vbCrLf & vbCrLf & _
             "このスクリプトは、Fortniteの設定ファイルを定期的に監視して" & vbCrLf & _
             "マウス加速が勝手にオン(False)になっていたら" & vbCrLf & _
             "強制的にオフ(True)へ書き換えてくれます。" & vbCrLf & vbCrLf & _
             "・スタートアップに登録すれば、PC起動後に自動で監視をスタートできます。" & vbCrLf & _
             "・終了したい場合は、検知時のダイアログで[いいえ]を押すか" & vbCrLf & _
             "  タスクマネージャから'wscript.exe'を終了してください。" & vbCrLf & vbCrLf & _
             "監視を今すぐ開始しますか？ (いいえ で中断)"

Dim startupRet
startupRet = MsgBox(startupMsg, vbYesNo + vbInformation, "起動確認")

If startupRet = vbNo Then
    MsgBox "了解しました。また必要になったら起動してくださいね。", vbInformation, "スクリプト終了"
    WScript.Quit
End If

'-------------------------------------------------------------------------
' ファイル存在チェック
'-------------------------------------------------------------------------
If Not fso.FileExists(configFilePath) Then
    MsgBox "おっと、『" & configFilePath & "』が見つかりませんでした！" & vbCrLf & _
           "「設定ファイルがそもそも無い」か、「フォルダの場所が違う」可能性があります。" & vbCrLf & vbCrLf & _
           "このままでは監視できないので、スクリプトを終了します。" & vbCrLf & _
           "ファイルの場所をもう一度ご確認ください！", _
           vbExclamation, _
           "ファイルが見つかりません"
    WScript.Quit
End If

'-------------------------------------------------------------------------
' メインループ開始
'-------------------------------------------------------------------------
Do While True
    
    On Error Resume Next  ' エラーが出てもスクリプトを継続できるようにする

    Dim fileObj, fileContent, accelFound
    accelFound = False
    
    '---------------------------------------------------------------------
    ' 設定ファイルを読み込み
    '---------------------------------------------------------------------
    Set fileObj = fso.OpenTextFile(configFilePath, 1) ' 1=ForReading
    fileContent = fileObj.ReadAll
    fileObj.Close
    
    '---------------------------------------------------------------------
    ' "bDisableMouseAcceleration=False" を探す
    '---------------------------------------------------------------------
    If InStr(1, fileContent, "bDisableMouseAcceleration=False", vbTextCompare) > 0 Then
        accelFound = True
    End If

    '---------------------------------------------------------------------
    ' 見つかったら強制的に True に書き換え & 通知
    '---------------------------------------------------------------------
    If accelFound = True Then
        
        fileContent = Replace(fileContent, _
                              "bDisableMouseAcceleration=False", _
                              "bDisableMouseAcceleration=True", 1, -1, vbTextCompare)
        
        ' 上書き保存
        Set fileObj = fso.OpenTextFile(configFilePath, 2)  ' 2=ForWriting
        fileObj.Write fileContent
        fileObj.Close
        
        ' 通知 (Yes/Noボタンで続行or終了を選択)
        Dim ret
        ret = MsgBox( _
            "ビビッ！ 不穏なマウス加速設定を発見しました！" & vbCrLf & _
            "そっとオフに書き換えておきましたのでご安心ください。" & vbCrLf & vbCrLf & _
            "Fortniteを再起動すれば設定が反映されます。" & vbCrLf & vbCrLf & _
            "まだまだ見張っておきますか？" & vbCrLf & _
            "(「いいえ」を押すとスクリプトを終了します)", _
            vbYesNo + vbInformation, _
            "Fortniteマウス加速 監視官")

        If ret = vbNo Then
            MsgBox "了解しました。監視業務からは撤退します！", _
                   vbInformation, _
                   "スクリプト終了"
            WScript.Quit
        End If
    End If
    
    '---------------------------------------------------------------------
    ' エラーハンドリングをリセット
    '---------------------------------------------------------------------
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました。(エラー番号=" & Err.Number & ") " & vbCrLf & _
               Err.Description, vbExclamation, "エラー通知"
    End If
    On Error GoTo 0  ' VBScriptのエラー処理モードを通常に戻す

    '---------------------------------------------------------------------
    ' 一定時間スリープ
    '---------------------------------------------------------------------
    WScript.Sleep SLEEP_INTERVAL

Loop