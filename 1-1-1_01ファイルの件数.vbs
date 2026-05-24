Option Explicit

'========================================
' フォルダー選択して
' ファイル件数を取得する
'========================================

Dim shell
Dim folderPath

'Shell.Application生成
Set shell = CreateObject("Shell.Application")

'フォルダー選択ダイアログ表示
Set folderPath = shell.BrowseForFolder( _
    0, _
    "フォルダーを選択してください", _
    0 _
)

'キャンセル時
If folderPath Is Nothing Then

    MsgBox "キャンセルされました"

    WScript.Quit

End If

'選択フォルダー取得
Dim targetPath

targetPath = folderPath.Self.Path

'----------------------------------------
' FileSystemObject生成
'----------------------------------------
Dim fso

Set fso = CreateObject( _
    "Scripting.FileSystemObject" _
)

'フォルダー取得
Dim folderObj

Set folderObj = fso.GetFolder( _
    targetPath _
)

'ファイル件数表示
MsgBox _
    "ファイル件数 : " & _
    CStr(folderObj.Files.Count)

'----------------------------------------
' 解放
'----------------------------------------
Set folderObj = Nothing
Set fso = Nothing
Set folderPath = Nothing
Set shell = Nothing