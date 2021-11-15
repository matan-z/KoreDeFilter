Attribute VB_Name = "m_FilteringByArray"
Option Explicit

Public i As Long, j As Long, Col As Long, selectedCol As Long, AFS As Long, AFE As Long
Public buf As String, adrs As String
Public isTable As Boolean
Public tableName As String

Sub KoreDeFilter_()
    f_FilterTool.Show vbModeless
End Sub

Sub KoreDeFilter_R(control As IRibbonControl)
    f_FilterTool.Show vbModeless
End Sub

'2
Sub KoreDeFilter(ByVal text As String)
    Dim lst As ListObject
    Set lst = ActiveCell.ListObject
    f_FilterTool.SelectedFile.Caption = ActiveWorkbook.Name
    If f_FilterTool.TextBox1 <> "" Then
        Dim myRng As Range
        adrs = ActiveCell.Address
        Dim ws As Worksheet
        Set ws = ActiveSheet
        Dim c As c_FilteringByArray: Set c = New c_FilteringByArray
        If lst Is Nothing Then
            If ws.AutoFilterMode Then
                Set myRng = ws.AutoFilter.Range
                Call c.FilteringByRange(myRng, f_FilterTool.TextBox1.Value) '2-1
            Else
                MsgBox "オートフィルタが設定されていません"
            End If
        Else
            Dim tableName As String
            tableName = ws.ListObjects(1).Name
            If ws.ListObjects(tableName).ShowAutoFilter Then
                Set myRng = ws.ListObjects(tableName).AutoFilter.Range
                Call c.FilteringByRange(myRng, f_FilterTool.TextBox1.Value)  '2-1
            Else
                MsgBox "オートフィルタが設定されていません"
            End If
        End If
    Else
        MsgBox "テキストボックスに何も入っていません"
    End If
End Sub

''【参考URL】https://desmondoshiwambo.wordpress.com/2012/02/23/how-to-copy-and-paste-texttofrom-clipboard-using-vba-microsoft-access/
'Sub CopyTextToClipboard(ByVal inText As String)
'    Dim objClipboard As Object
'    Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'    objClipboard.SetText inText
'    objClipboard.PutInClipboard
'    Set objClipboard = Nothing
'End Sub
'
''(does not require reference) :
'Function GetTextFromClipboard() As String
'    Dim objClipboard As Object
'    Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'    objClipboard.GetFromClipboard
'    On Error GoTo myError
'    GetTextFromClipboard = objClipboard.GetText
'    '******************************エラー処理
'myError:
'    If Err.Number <> 0 Then '2017/06/11に追記 エラーゼロで予期せぬエラーとなるため
'        Select Case Err.Number
'            Case -2147221404
'                MsgBox "クリップボードが空です" & vbCrLf & Err.Description, vbExclamation
'                'クリップボードでフィルタ.Show vbModeless
'            End
'            Case Else
'                On Error Resume Next
'                'クリップボードでフィルタ.Show vbModeless
'                MsgBox "予期せぬエラーが発生しました!", vbExclamation
'        End Select
'    End If
'    '******************************
'    Set objClipboard = Nothing
'End Function
