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

IF MsgBox(addInName & " �A�h�C�����C���X�g�[�����܂����H", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

Set objWshShell = CreateObject("WScript.Shell")

Set objFileSys = CreateObject("Scripting.FileSystemObject")

'�C���X�g�[����p�X�̍쐬
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'�t�@�C���R�s�[(�㏑��)
objFileSys.CopyFile  addInFileName ,installPath , True

Set objWshShell = Nothing
Set objFileSys = Nothing
'Excel �C���X�^���X��
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add

'�A�h�C���o�^
Set objAddin = objExcel.AddIns.Add(installPath, True)
objAddin.Installed = True

'Excel �I��
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing

IF Err.Number = 0 THEN
   MsgBox addInName & "�̃C���X�g�[�����I�����܂����B", vbInformation
ELSE
   MsgBox "�G���[���������܂����B" & vbCrLF & "���s�����m�F���Ă��������B", vbExclamation
End IF