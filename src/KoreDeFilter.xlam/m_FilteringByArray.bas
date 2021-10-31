Attribute VB_Name = "m_FilteringByArray"
Option Explicit

Public i As Long, j As Long, Col As Long, 最終行 As Long, selectedCol As Long, AFS As Long, AFE As Long
Public buf As String, adrs As String
Public isTable As Boolean
Public tableName As String

Sub KoreDeFilterTest()
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

'【参考URL】https://desmondoshiwambo.wordpress.com/2012/02/23/how-to-copy-and-paste-texttofrom-clipboard-using-vba-microsoft-access/
Sub CopyTextToClipboard(ByVal inText As String)
    Dim objClipboard As Object
    Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    objClipboard.SetText inText
    objClipboard.PutInClipboard
    Set objClipboard = Nothing
End Sub

Sub テキストボックスでカンマ区切り(ByVal text As String)
    Dim CB As Variant
    Dim クリップ() As String
    クリップ = 改行で格納(text)
    Dim クリップ配列() As String, i As Long
    ReDim クリップ配列(UBound(クリップ))
    For i = 0 To UBound(クリップ)
        If クリップ(i) <> "" Then
            クリップ配列(i) = Replace(クリップ(i), vbCr, "")
        End If
    Next i
    '【参考URL】https://oshiete.goo.ne.jp/qa/6871211.html
    Dim c As Variant
    Dim 個数 As Long
    Dim 配列 As Variant
    Set DB = CreateObject("Scripting.Dictionary")
    For Each c In クリップ配列
        DB(c) = 1
    Next c
    個数 = DB.Count '参考、新配列の個数
    配列 = DB.Keys '新配列の展開
    buf = ""
    For i = 0 To UBound(配列)
        buf = buf + 配列(i) & ","
    Next i
    If Len(buf) > 0 Then
        buf = Left(buf, Len(buf) - 1)
    End If
    CopyTextToClipboard (buf)
    MsgBox buf & vbLf & "をクリップボードに置きました"
End Sub
Sub 空白削除(ByVal text As String)
    Dim CB As Variant
    Dim クリップ() As String
    クリップ = 改行で格納(text)
    Dim クリップ配列() As String, i As Long
    ReDim クリップ配列(UBound(クリップ))
    For i = 0 To UBound(クリップ)
        クリップ配列(i) = Replace(クリップ(i), vbCr, "")
    Next i
    '重複削除
    '【参考URL】https://oshiete.goo.ne.jp/qa/6871211.html
    Dim c As Variant
    Dim 個数 As Long
    Dim 配列 As Variant
    Set DB = CreateObject("Scripting.Dictionary")
    For Each c In クリップ配列
        DB(c) = 1
    Next c
    個数 = DB.Count
    配列 = DB.Keys
    i = 0
    j = 0
    buf = ""
    For i = 0 To UBound(配列)
        If 配列(i) <> "" And i <> UBound(配列) Then
            buf = buf & 配列(i) & vbLf
            j = j + 1
        Else
            If 配列(i) <> "" And i = UBound(配列) Then
                buf = buf & 配列(i)
                j = j + 1
            End If
        End If
    Next
    f_FilterTool.TextBox1.Value = ""
    f_FilterTool.TextBox1.Value = buf
End Sub

Function 改行で格納(str As String) As String()
    改行で格納 = Split(TrimLF(str), vbLf)
End Function
Sub 複数セル選択の場合()
    Dim ad, buf As String
    ad = Selection.Address
    Dim adArray() As String
    adArray = Split(ad, ",")
    Dim 配列 As Variant
    ReDim 配列(UBound(adArray))
    buf = ""
    For i = 0 To UBound(adArray)
        配列(i) = Range(adArray(i)).Value
        If i <> UBound(adArray) Then
            buf = buf + Range(adArray(i)).Value '+ ","
        Else
            buf = buf + Range(adArray(i)).Value
        End If
    Next i
    With New MSForms.DataObject
        .SetText buf '変数の値をDataObjectに格納する
        .PutInClipboard 'DataObjectのデータをクリップボードに格納する
    End With
End Sub
'【参考URL】https://desmondoshiwambo.wordpress.com/2012/02/23/how-to-copy-and-paste-texttofrom-clipboard-using-vba-microsoft-access/

'(does not require reference) :
Function GetTextFromClipboard() As String
    Dim objClipboard As Object
    Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    objClipboard.GetFromClipboard
    On Error GoTo myError
    GetTextFromClipboard = objClipboard.GetText
    '******************************エラー処理
myError:
    If Err.Number <> 0 Then '2017/06/11に追記 エラーゼロで予期せぬエラーとなるため
        Select Case Err.Number
            Case -2147221404
                MsgBox "クリップボードが空です" & vbCrLf & Err.Description, vbExclamation
                'クリップボードでフィルタ.Show vbModeless
            End
            Case Else
                On Error Resume Next
                'クリップボードでフィルタ.Show vbModeless
                MsgBox "予期せぬエラーが発生しました!", vbExclamation
        End Select
    End If
    '******************************
    Set objClipboard = Nothing
End Function
