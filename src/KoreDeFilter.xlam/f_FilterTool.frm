VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_FilterTool 
   Caption         =   "FilteringTool"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3030
   OleObjectBlob   =   "f_FilterTool.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "f_FilterTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim text As String
Dim i, j, k As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOSIZE As Long = &H1&
Private Const SWP_NOMOVE As Long = &H2&

Private Sub cbn_DeleteDuplicateAndBlank_Click()
    Dim c As c_FilteringByArray: Set c = New c_FilteringByArray
    
    Dim arr() As String
    arr = c.ArrByLf(TextBox1.Value)
    arr = c.VbCrRemoveArr(arr)
    Dim arr2 As Variant
    arr2 = c.DelDupe(arr)
    Dim i As Long
    Dim buf As String
    For i = LBound(arr2) To UBound(arr2)
        If Len(arr2(i)) > 0 Then
            buf = buf & arr2(i) & vbCrLf
        End If
    Next i
    TextBox1.Value = buf
End Sub
'1
Private Sub cbn_FilterByTextBox_Click()
    m_FilteringByArray.KoreDeFilter TextBox1.Value
End Sub

Private Sub CloseButton_Click() 'Closeボタン
    Unload Me
End Sub

Private Sub cbn_RemoveFilter_Click() 'フィルタ解除
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
End Sub

Private Sub TextBox1_Change()
    ThisWorkbook.Worksheets("Sheet1").Range("A1").Value = TextBox1.Value
End Sub

Private Sub UserForm_Initialize()
    Me.lbl_ThisName.Caption = ThisWorkbook.Name
    Me.Height = 404.25
    Me.Width = 163.5
    With TextBox1
        .MultiLine = True
        .EnterKeyBehavior = True
        .Value = Sheet1.Range("A1")
    End With
End Sub

Private Sub UserForm_Terminate()
    Sheet1.Range("A1") = f_FilterTool.TextBox1.Value
End Sub
