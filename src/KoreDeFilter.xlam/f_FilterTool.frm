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

Private Sub ComboBox1_Change()
    Dim AcBook As String
    Dim 配列() As Variant
    ReDim 配列(Workbooks.Count - 2)
    Dim wbook As Workbook
    For j = 0 To ComboBox1.ListCount - 1
        For i = 0 To Workbooks.Count - 2
            For Each wbook In Workbooks
                If wbook.Name <> ThisWorkbook.Name Then
                    If wbook.Name = ComboBox1.List(j) Then
                        k = k + 1
                    End If
                End If
            Next wbook
            If k < 1 Then
                ComboBox1.List(j).Remove
            End If
        Next i
    Next j
    If ComboBox1.Value <> ThisWorkbook.Name And ComboBox1 <> "" Then
        AcBook = ComboBox1.Value
        Workbooks(AcBook).Activate
    Else
        MsgBox "プルダウンでファイルを選択してください"
    End If
End Sub
Private Sub CommandButton1_Click()
    でフィルタとカンマ区切る.でフィルタ TextBox1.Value
End Sub
Private Sub CommandButton2_Click()
    m_FilteringByArray.空白削除 TextBox1.Value
End Sub
'1
Private Sub FilterButton_Click()
    m_FilteringByArray.KoreDeFilter TextBox1.Value
End Sub
Private Sub CurrentFileGet_Click() '最新ファイル情報取得
    Dim i As Long
    If Workbooks.Count - 2 > -1 Then
        Dim 配列() As Variant
        ReDim 配列(Workbooks.Count - 2)
        Dim wbook As Workbook
        For i = 0 To Workbooks.Count - 2
            For Each wbook In Workbooks
                If wbook.Name <> ThisWorkbook.Name Then
                    配列(i) = wbook.Name
                    i = i + 1
                End If
            Next wbook
        Next i
        ComboBox1.List = arr
    Else
        MsgBox "対象ファイルがありません"
    End If
End Sub

Private Sub ActivateButton_Click()
    Dim AcBook As String
    If ComboBox1.Value <> "" Then
        Set フィルター対象 = Workbooks(ComboBox1.Value)
        フィルター対象.Activate
    Else
        MsgBox "プルダウンでファイルを選択してください"
    End If
End Sub

Private Sub CloseButton_Click() 'Closeボタン
    Unload Me
End Sub

Private Sub HaishiFilterButton_Click()
    Call m_f_Filtering.Haishi
End Sub
Private Sub RemoveFilter_Click() 'フィルタ解除
    ' If ComboBox1.Value <> "" Then
    ' Set フィルター対象 = Workbooks(ComboBox1.Value)
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    ' SelectedFile.Caption = ActiveWorkbook.Name
    ' Set フィルター対象 = Workbooks(ActiveWorkbook.Name)
    '
    ' Dim c As c_func: Set c = New c_func
    '
    ' Dim tableFileter As Boolean
    
    ' tableFileter = c.FilterIsOn(ActiveSheet.ListObjects(1))
    '
    ' If ActiveWorkbook.Name <> フィルター対象.Name Then フィルター対象.Activate
    '' Application.Wait [Now() + "00:00:00.5"]
    ' If フィルター対象.ActiveSheet.AutoFilterMode Then
    ' With ActiveSheet
    ' If .FilterMode Then .ShowAllData
    ' End With
    ' End If
    '' End If
End Sub
Private Sub BlankDeleteButton_Click()
    m_FilteringByArray.空白削除 TextBox1.Value
End Sub
Private Sub CommandButton3_Click()
    Dim c As Range
    For Each c In Selection
        Dim cellValue As String
        cellValue = c.Value
        If Left(cellValue, 1) <> "■" Then
            cellValue = "■" & cellValue
        End If
        cellValue = StrConv(cellValue, vbWide)
        c.Value = cellValue
    Next c
End Sub
Private Sub CommandButton4_Click()
    Dim c As Range
    For Each c In Selection
        Dim cellValue As String
        cellValue = c.Value
        If Left(cellValue, 1) = "■" Then
            cellValue = Replace(cellValue, "■", "")
        End If
        cellValue = StrConv(cellValue, vbNarrow)
        c.Value = cellValue
    Next c
End Sub

Private Sub ShinkiFilterButton_Click()
    Call m_f_Filtering.Shinki
End Sub

Private Sub TextBox1_Change()
    ThisWorkbook.Worksheets("Sheet1").Range("A1").Value = TextBox1.Value
End Sub
'Private Sub f_FilterTool_Activate()
'Call SetWindowPos(GetForegroundWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE OrSWP_NOSIZE)
' SelectedFile.Caption = ActiveWorkbook.Name
'
' If ThisWorkbook.Worksheets("Sheet1").Range("A1").Value <> "" Then
' TextBox1.Value = ThisWorkbook.Worksheets("Sheet1").Range("A1").Value
' End If
'
' If Workbooks.Count > 0 Then
'
' Dim arr() As Variant
' ReDim arr(Workbooks.Count)
' Dim wbook As Workbook
' For i = 0 To Workbooks.Count
' For Each wbook In Workbooks
' If wbook.Name <> ThisWorkbook.Name Then
' arr(i) = wbook.Name
' i = i + 1

' End If
' Next wbook
' Next i
' ComboBox1.List = arr
' Else
'' MsgBox "対象ファイルがありません"
' End If
'End Sub
Private Sub UserForm_Initialize()
    Me.Label1.Caption = ThisWorkbook.Name
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
