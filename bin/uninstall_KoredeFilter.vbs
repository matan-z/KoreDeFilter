'https://www.aruse.net/entry/2018/09/13/081734
On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin

'�A�h�C������ݒ�
addInName = "KoreDeFilter"
addInFileName = "KoreDeFilter.xlam" 

IF MsgBox(addInName & "���A���C���X�g�[�����܂����H", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

'Excel �C���X�^���X��
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add

'�A�h�C���o�^����
For i = 1 To objExcel.Addins.Count
  Set objAddin = objExcel.Addins.item(i)
  If objAddin.Name = addInFileName Then
    objAddin.Installed = False
  End If
Next

'Excel �I��
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing

IF Err.Number = 0 THEN
   MsgBox addInName &"�̃A���C���X�g�[�����I�����܂����B", vbInformation
ELSE
   MsgBox "�G���[���������܂����B" & vbCrLF & "���s�����m�F���Ă��������B", vbExclamation
End IF
