Option Explicit  ' �ϐ��錾�R���h��

'-------------------------------------------------------------------------
' �ϐ��錾
'-------------------------------------------------------------------------
Dim fso              ' �t�@�C������p (Scripting.FileSystemObject)
Dim shell            ' �V�F������p (WScript.Shell)
Dim localAppData     ' %localappdata% �̃p�X�W�J�p
Dim configFilePath   ' Fortnite�ݒ�t�@�C���̃p�X
Dim SLEEP_INTERVAL   ' �Ď��Ԋu(�~���b)

SLEEP_INTERVAL = 30000 ' 30�b�����Ƀ`�F�b�N(1000=1�b, 30000=30�b)

'-------------------------------------------------------------------------
' �I�u�W�F�N�g�쐬
'-------------------------------------------------------------------------
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

'-------------------------------------------------------------------------
' �t�H�[�g�i�C�g�ݒ�t�@�C���̃p�X��g�ݗ��Ă�
'   (��) %localappdata%\FortniteGame\Saved\Config\WindowsClient\GameUserSettings.ini
'-------------------------------------------------------------------------
localAppData   = shell.ExpandEnvironmentStrings("%localappdata%") ' ���ϐ���W�J
configFilePath = localAppData & "\FortniteGame\Saved\Config\WindowsClient\GameUserSettings.ini"

'-------------------------------------------------------------------------
' �t�@�C�����݃`�F�b�N
'   �����t�@�C����������΁A���b�Z�[�W���o���đ��I��
'-------------------------------------------------------------------------
If Not fso.FileExists(configFilePath) Then
    MsgBox "�����ƁA�w" & configFilePath & "�x��������܂���ł����I" & vbCrLf & _
           "�u�ݒ�t�@�C�����������������v���A�u�t�H���_�̏ꏊ���Ⴄ�v�\��������܂��B" & vbCrLf & vbCrLf & _
           "���̂܂܂ł͊Ď��ł��Ȃ��̂ŁA�X�N���v�g���I�����܂��B" & vbCrLf & _
           "�t�@�C���̏ꏊ��������x���m�F���������I", _
           vbExclamation, _
           "�t�@�C����������܂���"
    WScript.Quit
End If

'-------------------------------------------------------------------------
' ���C�����[�v�J�n
'-------------------------------------------------------------------------
Do While True
    
    On Error Resume Next  ' �G���[���o�Ă��X�N���v�g���p���ł���悤�ɂ���

    Dim fileObj, fileContent, accelFound
    accelFound = False
    
    '---------------------------------------------------------------------
    ' �ݒ�t�@�C����ǂݍ���
    '---------------------------------------------------------------------
    Set fileObj = fso.OpenTextFile(configFilePath, 1) ' 1=ForReading
    fileContent = fileObj.ReadAll
    fileObj.Close
    
    '---------------------------------------------------------------------
    ' "bDisableMouseAcceleration=False" ��T��
    '---------------------------------------------------------------------
    If InStr(1, fileContent, "bDisableMouseAcceleration=False", vbTextCompare) > 0 Then
        accelFound = True
    End If

    '---------------------------------------------------------------------
    ' ���������狭���I�� True �ɏ������� & �ʒm
    '---------------------------------------------------------------------
    If accelFound = True Then
        
        fileContent = Replace(fileContent, _
                              "bDisableMouseAcceleration=False", _
                              "bDisableMouseAcceleration=True", 1, -1, vbTextCompare)
        
        ' �㏑���ۑ�
        Set fileObj = fso.OpenTextFile(configFilePath, 2)  ' 2=ForWriting
        fileObj.Write fileContent
        fileObj.Close
        
        ' �ʒm (Yes/No�{�^���ő��sor�I����I��)
        Dim ret
        ret = MsgBox( _
            "�r�r�b�I �s���ȃ}�E�X�����ݒ�𔭌����܂����I" & vbCrLf & _
            "�����ƃI�t�ɏ��������Ă����܂����̂ł����S���������B" & vbCrLf & vbCrLf & _
            "Fortnite���ċN������ΐݒ肪���f����܂��B" & vbCrLf & vbCrLf & _
            "�܂��܂��������Ă����܂����H" & vbCrLf & _
            "(�u�������v�������ƃX�N���v�g���I�����܂�)", _
            vbYesNo + vbInformation, _
            "Fortnite�}�E�X���� �Ď���")

        If ret = vbNo Then
            MsgBox "�������܂����B�Ď��Ɩ�����͓P�ނ��܂��I", _
                   vbInformation, _
                   "�X�N���v�g�I��"
            WScript.Quit
        End If
    End If
    
    '---------------------------------------------------------------------
    ' �G���[�n���h�����O�����Z�b�g
    '---------------------------------------------------------------------
    If Err.Number <> 0 Then
        MsgBox "�G���[���������܂����B(�G���[�ԍ�=" & Err.Number & ") " & vbCrLf & _
               Err.Description, vbExclamation, "�G���[�ʒm"
    End If
    On Error GoTo 0  ' VBScript�̃G���[�������[�h��ʏ�ɖ߂�

    '---------------------------------------------------------------------
    ' ��莞�ԃX���[�v
    '---------------------------------------------------------------------
    WScript.Sleep SLEEP_INTERVAL

Loop
