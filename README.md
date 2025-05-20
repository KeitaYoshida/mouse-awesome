# 1. 目的

本スクリプトは、Fortniteの設定ファイル (GameUserSettings.ini) において
「bDisableMouseAcceleration=False」になってしまう問題を自動的に検知し
強制的にオフ(True)へ書き換える処理を行うためのものです。

* Fortnite側の仕様やアップデートにより、マウス加速が勝手にオンになるケースを防止し、安定したマウス操作を確保します。
* 完全に裏側で動作し、加速を検知した場合のみ通知してくれます。

## 本スクリプト（VBScript）を採用理由
	1. Windows標準の機能(WSH) で動くため、インストール不要。
	2. コードがテキスト形式でオープンに公開でき、ユーザーが内容をチェックしやすい。
	3. コンソール画面を表示せずに「裏で常駐」できる。

⸻

# 2. インストール方法
mouse-awesome.vbsをダウンロード

## 2.1. 自分でスクリプトを作成する場合
	1.メモ帳(Notepad)を起動
	　Windowsキー + R を押し、「notepad」と入力してOKします。
	2.スクリプトを貼り付け
	　このドキュメントの末尾にあるVBScriptコードをコピー＆ペーストしてください。
	3.文字コードを指定して保存
	　•「ファイル > 名前を付けて保存」を選択し、
	　•ファイルの種類を「すべてのファイル(*.*)」にし、拡張子を「.vbs」にして保存します。
	　•文字コード(エンコード) は「ANSI」または「Unicode(UTF-16)」を選ぶのがおすすめです。
	　•「UTF-8(BOMなし)」は文字化けやエラーの原因になることがあります。
	4.スクリプトを実行してテスト
	　•作成した .vbs ファイルをダブルクリックすると、裏で実行されます。
	　•もしFortniteの設定ファイル GameUserSettings.ini が見つからなかった場合は、エラー通知を出して即終了します。
	　•設定ファイルが存在し、かつ「bDisableMouseAcceleration=False」が検出された時だけ通知が出ます。

## 2.2. 自動起動（スタートアップ）へ登録する方法
	•方法A: スタートアップフォルダにショートカットを置く
	　1.「Windowsキー + R」を押し、shell:startup と入力 → [Enter]。
	　2.表示されたフォルダに、作成した .vbs ファイルの「ショートカット」を配置します。
	　3.Windows起動時（正確にはユーザーログイン時）にこのスクリプトが自動で走り、監視を開始します。
	•方法B: レジストリのRunキーを使う
	　1.レジストリエディタを起動し、以下のキーを開きます。

```
HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
```
	　2.文字列値(REG_SZ)を新規作成し、名前を例として FortniteMouseAccelWatcher 等に設定。
	　3.データに WScript.exe "C:\Path\To\FortniteMouseAccelCheck.vbs" と記入。
	　4.次回ログイン時に自動で実行されるようになります。

# 3. 設計・危険性への配慮
	1.オープンなテキスト形式
	　•vbs はテキストファイルなので、誰でも中身をメモ帳で閲覧できます。
	　•配布されるスクリプトが怪しい操作（ファイル削除など）を行っていないか、確認しやすいです。
	2.自分で作るのが安心
	　•第三者がコンパイルした実行ファイル(.exe)を使うと、中で何をしているか分かりにくいというリスクがあります。
	　•そこで、自らの手で .vbs を作成し、中身を把握した上で動かすことが安全性向上に繋がります。
	3.バッチファイル(.bat) との違い
	　•.bat はコンソールを表示しっぱなしになるため、見た目が煩わしくなるケースがあります。
	　•.vbs は コンソール非表示で「裏で常駐」し、必要時だけMsgBoxで通知を行うので、簡易デーモン的な働きをさせやすいのが利点です。
	4.スクリプトの終了方法
	　•本スクリプトは標準だとタスクトレイにアイコンを表示しないため、手動で終了する場合はタスクマネージャ→wscript.exe プロセスを終了 させる必要があります。
	　•なお、加速を検知した際に表示するメッセージボックスが、[いいえ]を押すとスクリプトを終了する設計になっています。
	　•万一、延々と動き続けると困る場合は、コメントなどを編集して独自の終了ダイアログを追加してください。

⸻

# 4. その他の留意事項
	•本スクリプトでは「Fortniteの設定ファイルパスが決め打ち」になっています。
	•多くの環境で %localappdata%\FortniteGame\Saved\Config\WindowsClient\GameUserSettings.ini が使用される想定ですが、
	•万一フォルダ構成が変わると検知できなくなる場合があります。
	•スクリプト内の文字列置換は "bDisableMouseAcceleration=False" → "bDisableMouseAcceleration=True" のみ実施しています。
	•Fortnite側の仕様変更等で設定キー名が変化した場合には、スクリプトも修正してください。
	•30秒ごとにファイルをチェックする設計です（SLEEP_INTERVAL = 30000）。
	•もしPC負荷を下げたい・あるいは更に頻度を上げたい場合は、この値を適宜変更してください。

⸻

# 5. VBScriptコード全文

以下が実際のスクリプト（.vbs）例です。テキストエディタでコピー＆ペーストし、先述の通り「ANSI」または「Unicode(UTF-16)」で保存してください。
``` vbs
Option Explicit  ' 変数宣言漏れを防ぐ

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
'   (例) %localappdata%\FortniteGame\Saved\Config\WindowsClient\GameUserSettings.ini
'-------------------------------------------------------------------------
localAppData   = shell.ExpandEnvironmentStrings("%localappdata%") ' 環境変数を展開
configFilePath = localAppData & "\FortniteGame\Saved\Config\WindowsClient\GameUserSettings.ini"

'-------------------------------------------------------------------------
' ファイル存在チェック
'   もしファイルが無ければ、メッセージを出して即終了
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
```
