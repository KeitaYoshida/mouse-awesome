Option Explicit  ' 変数宣言漏れを防ぐ

'-------------------------------------------------------------------------
' このスクリプトの概要
'   - Fortnite の設定ファイル (GameUserSettings.ini) を定期的に監視し、
'     bDisableMouseAcceleration=False を見つけたら True に書き換える。
'   - 初回起動時のみ、簡単な案内メッセージを表示。
'   - 2回目以降は通知を一切出さずに自動修正を行う(サイレントモード)。
'   - 終了するには、タスクマネージャで wscript.exe を終了させる。
'-------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------
' 変数宣言
'-------------------------------------------------------------------------
Dim fso               ' ファイル操作用 (Scripting.FileSystemObject)
Dim shell             ' シェル操作用 (WScript.Shell)
Dim localAppData      ' %localappdata% のパス展開用
Dim configFilePath    ' Fortnite設定ファイルのパス
Dim SLEEP_INTERVAL    ' 監視間隔(ミリ秒)
Dim markerFilePath    ' 初回起動を記録するためのマーカーファイル

SLEEP_INTERVAL = 30000 ' 30秒おきにチェック(1000=1秒, 30000=30秒)

'-------------------------------------------------------------------------
' オブジェクト作成
'-------------------------------------------------------------------------
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

'-------------------------------------------------------------------------
' Fortnite設定ファイルのパスを組み立てる
'-------------------------------------------------------------------------
localAppData   = shell.ExpandEnvironmentStrings("%localappdata%")
configFilePath = localAppData & "\FortniteGame\Saved\Config\WindowsClient\GameUserSettings.ini"

'-------------------------------------------------------------------------
' 初回起動判定用ファイルのパス(スクリプトと同じフォルダに作る想定)
'   ※必ずしも同じフォルダでなくてもOK。任意の場所で管理可能です。
'-------------------------------------------------------------------------
markerFilePath = fso.GetParentFolderName(WScript.ScriptFullName) & _
                 "\MouseAccelWatcher_FirstRunDone.txt"

Dim isFirstRun
isFirstRun = False
If Not fso.FileExists(markerFilePath) Then
    '---------------------------------------------------------------------
    ' マーカーファイルがまだ無い → 今回が初回起動
    '---------------------------------------------------------------------
    isFirstRun = True
    
    ' 初回起動案内メッセージ
    Dim msg
    msg = "【Fortniteマウス加速ウォッチャー - 初回起動】" & vbCrLf & vbCrLf & _
          "このスクリプトは、Fortniteの設定ファイルを30秒ごとに監視し" & vbCrLf & _
          "マウス加速がオン(False)になっていたら、こっそりオフ(True)に" & vbCrLf & _
          "書き換えてくれるツールです。" & vbCrLf & vbCrLf & _
          "◆ 終了したいときは、タスクマネージャから『wscript.exe』を終了してください。" & vbCrLf & vbCrLf & _
          "このまま監視を開始します。"
    
    MsgBox msg, vbInformation, "初回起動のお知らせ"

    ' マーカーファイルを作成（中身は特に不要）
    Dim markerFile
    Set markerFile = fso.CreateTextFile(markerFilePath, True)
    markerFile.WriteLine "This file indicates the script has run at least once."
    markerFile.Close
End If

'-------------------------------------------------------------------------
' Fortnite設定ファイルが存在しない場合は終了
'-------------------------------------------------------------------------
If Not fso.FileExists(configFilePath) Then
    ' 初回/2回目以降に関係なく、ファイルが無ければ終わる
    Dim errMsg
    errMsg = "おっと、『" & configFilePath & "』が見つかりませんでした！" & vbCrLf & _
             "フォートナイトの設定ファイルが無い、または場所が違う可能性があります。" & vbCrLf & vbCrLf & _
             "このままでは監視できないので、スクリプトを終了します。"
    If isFirstRun Then
        MsgBox errMsg, vbExclamation, "ファイルが見つかりません"
    Else
        ' 2回目以降なら、通知も不要かもしれませんが、
        ' 何が起こったか分からなくなるので、一応エラー表示しておく事を推奨。
        MsgBox errMsg, vbExclamation, "ファイルが見つかりません"
    End If
    WScript.Quit
End If

'-------------------------------------------------------------------------
' メインループ開始 (通知は一切出さないサイレントモード)
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
    ' 見つかったら強制的に True に書き換える (サイレントに実施)
    '---------------------------------------------------------------------
    If accelFound = True Then
        fileContent = Replace(fileContent, _
                              "bDisableMouseAcceleration=False", _
                              "bDisableMouseAcceleration=True", 1, -1, vbTextCompare)
        
        ' 上書き保存
        Set fileObj = fso.OpenTextFile(configFilePath, 2) ' 2=ForWriting
        fileObj.Write fileContent
        fileObj.Close
        ' 通知や終了確認は一切行わない
    End If
    
    '---------------------------------------------------------------------
    ' エラーハンドリングをリセット (何かあれば一度だけ表示)
    '---------------------------------------------------------------------
    If Err.Number <> 0 Then
        ' 万一ここで何らかのエラーが発生した場合、初回かどうかに関わらず一回出す
        ' (それでも良い、という場合はこの MsgBox も消して完全サイレントにしても可)
        MsgBox "エラーが発生しました。(エラー番号=" & Err.Number & ") " & vbCrLf & _
               Err.Description, vbExclamation, "エラー通知"
    End If
    On Error GoTo 0  ' VBScriptのエラー処理モードを通常に戻す

    '---------------------------------------------------------------------
    ' 一定時間スリープ
    '---------------------------------------------------------------------
    WScript.Sleep SLEEP_INTERVAL

Loop