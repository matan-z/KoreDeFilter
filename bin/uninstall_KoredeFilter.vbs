'https://www.aruse.net/entry/2018/09/13/081734
On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin

'アドイン情報を設定
addInName = "KoreDeFilter"
addInFileName = "KoreDeFilter.xlam" 

IF MsgBox(addInName & "をアンインストールしますか？", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

'Excel インスタンス化
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add

'アドイン登録解除
For i = 1 To objExcel.Addins.Count
  Set objAddin = objExcel.Addins.item(i)
  If objAddin.Name = addInFileName Then
    objAddin.Installed = False
  End If
Next

'Excel 終了
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing

IF Err.Number = 0 THEN
   MsgBox addInName &"のアンインストールが終了しました。", vbInformation
ELSE
   MsgBox "エラーが発生しました。" & vbCrLF & "実行環境を確認してください。", vbExclamation
End IF
