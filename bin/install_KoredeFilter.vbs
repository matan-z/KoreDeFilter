'https://www.aruse.net/entry/2018/09/13/081734

Dim installPath
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin

'アドイン名設定
addInName = "KoreDeFilter"
addInFileName = "KoreDeFilter.xlam"

IF MsgBox(addInName & " アドインをインストールしますか？", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'スクリプトのフォルダからソースパスを作成
Dim scriptDir
scriptDir = objFileSys.GetParentFolderName(WScript.ScriptFullName)
Dim sourcePath
sourcePath = objFileSys.BuildPath(scriptDir, addInFileName)

'ソースファイルの存在確認
If Not objFileSys.FileExists(sourcePath) Then
  MsgBox "アドインファイルが見つかりません。" & vbCrLf & sourcePath, vbExclamation
  WScript.Quit
End If

'インストール先パスの作成
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
Dim addInsDir
addInsDir = objWshShell.SpecialFolders("Appdata") & "\Microsoft\AddIns"
If Not objFileSys.FolderExists(addInsDir) Then
  objFileSys.CreateFolder addInsDir
End If
installPath = addInsDir & "\" & addInFileName

'ファイルコピー(上書き)
On Error Resume Next
objFileSys.CopyFile sourcePath, installPath, True
If Err.Number <> 0 Then
  MsgBox "ファイルコピーでエラーが発生しました。" & vbCrLf & "エラー: " & Err.Number & " " & Err.Description, vbExclamation
  WScript.Quit
End If
On Error GoTo 0

Set objWshShell = Nothing
Set objFileSys = Nothing

'Excel インスタンス化
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
  MsgBox "Excelの起動でエラーが発生しました。" & vbCrLf & "エラー: " & Err.Number & " " & Err.Description, vbExclamation
  WScript.Quit
End If
objExcel.Workbooks.Add

'アドイン登録
Set objAddin = objExcel.AddIns.Add(installPath, True)
If Err.Number <> 0 Then
  MsgBox "アドイン登録でエラーが発生しました。" & vbCrLf & "エラー: " & Err.Number & " " & Err.Description, vbExclamation
  objExcel.Quit
  WScript.Quit
End If
objAddin.Installed = True

'Excel 終了
objExcel.Quit
On Error GoTo 0

Set objAddin = Nothing
Set objExcel = Nothing

MsgBox addInName & "のインストールが終了しました。", vbInformation
