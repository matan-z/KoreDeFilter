VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_FilteringByArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'3 '2-1
Function FilteringByRange(myRng As Range, text As String)
    AFS = myRng.Column '
    AFE = myRng.Columns.Count '列が何列あるか

    Col = ActiveCell.Column '現在の選択セルの列
    Dim arr() As String
    text = Trim(text)
    Dim c1 As c_FilteringByArray: Set c1 = New c_FilteringByArray
    arr = c1.ArrByLf(text) '3-1
    Dim arr2() As String
    arr2 = c1.VbCrRemoveArr(arr) '3-2
    Dim arr3 As Variant
    arr3 = Me.DelDupe(arr2)
    Call Me.Filter(ActiveWorkbook, arr3) '3-3
End Function
'4 '3-1
Function ArrByLf(str As String) As String()
    ArrByLf = Split(Me.TrimArr(str), vbLf) '4-1
End Function
'http://www.relief.jp/docs/013815.html ▼文字列の前と後ろ・セルの先頭と末尾の改行を削除する
'5 '4-1
Function TrimArr(str As String) As String
    Dim strTmp As String
    strTmp = str
    Do Until Left(strTmp, 1) <> vbLf
        strTmp = Mid(strTmp, 2)
    Loop
    Do Until Right(strTmp, 1) <> vbLf
        strTmp = Left(strTmp, Len(strTmp) - 1)
    Loop
    TrimArr = strTmp
End Function

'6  '3-2
Function VbCrRemoveArr(arr() As String) As String()
    Dim textBoxArr() As String, i As Long
    ReDim textBoxArr(UBound(arr))
    '各vbCrを削除
    For i = 0 To UBound(arr)
        textBoxArr(i) = Trim(arr(i))
        textBoxArr(i) = Replace(arr(i), vbCr, "")
    Next i
    VbCrRemoveArr = textBoxArr
End Function

'7 '3-3
Function DelDupe(arr() As String) As Variant
    Dim delDupeArr As Variant
    Dim c As Variant
    Dim cnt As Long
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    For Each c In arr
        dic(c) = 0
    Next c
    cnt = dic.Count '参考、新配列の個数
    delDupeArr = dic.Keys '新配列の展開
    i = 0
    j = 0
    Dim buf As String
    buf = ""
    ' f_FilterTool.TextBox1.Value = ""
    For i = 0 To UBound(delDupeArr)
        If delDupeArr(i) <> "" And i <> UBound(delDupeArr) Then
            buf = buf & delDupeArr(i) & vbLf
            j = j + 1
        Else
            If delDupeArr(i) <> "" And i = UBound(delDupeArr) Then
                buf = buf & delDupeArr(i)
                j = j + 1
            End If
        End If
        f_FilterTool.TextBox1.Value = buf
    Next
    DelDupe = delDupeArr
End Function

'8
Function Filter(targetWb As Workbook, arr As Variant)
    Dim msg As String
    Dim all As String
    msg = ""
    Dim ws As Worksheet
    Set ws = targetWb.ActiveSheet
    ' If Col <= AFS Or Col > AFS + AFE Then
    selectedCol = (Col - AFS) + 1
    ' End If
    With ws
        ' On Error Resume Next
        If isTable = False Then
            .Cells(.AutoFilter.Range.Row, .AutoFilter.Range.Column).AutoFilter Field:=selectedCol, _
            Criteria1:=Array(arr), _
            Operator:=xlFilterValues
        Else
            .ListObjects(tableName).Range.AutoFilter Field:=selectedCol, _
            Criteria1:=Array(arr), _
            Operator:=xlFilterValues
        End If
        
        If WorksheetFunction.Subtotal(3, .Columns(selectedCol)) = 1 Then  'フィルター対象の列にデータがひとつしかなかったら
            MsgBox "該当データがないのでフィルターを戻します"
            ' ActiveSheet.Columuns(selectedCol ).ShowAllData
            '列のみフィルター 【参考】https://matome.naver.jp/odai/2143208158655847001
            .Cells(ActiveSheet.AutoFilter.Range.Row, ActiveSheet.AutoFilter.Range.Column).AutoFilter Field:=selectedCol
        Else
            i = 0
            j = 0
            buf = ""
            ' UserForm4.Show vbModeless
            For i = 0 To UBound(arr)
                Dim Obj As Range
                Set Obj = ws.Columns(Col).Find(What:=arr(i), LookIn:=xlValues, Lookat:=xlPart, SearchOrder:=xlByColumns)
                Dim found As String
                Dim notFound As String
                Dim both As String
                If Obj Is Nothing Then
                    buf = buf & arr(i) & vbCrLf
                    notFound = notFound & arr(i) & vbCrLf
                    both = both & i + 1 & " " & "×" & " " & arr(i) & vbCrLf
                Else
                    found = found & arr(i) & vbCrLf
                    both = both & i + 1 & " " & "〇" & " " & arr(i) & vbCrLf
                    j = j + 1
                End If
            Next i
            If buf <> "" Then
                buf = Mid(buf, 1, Len(buf) - 1)
            End If
            If buf = "" Then
                Dim WSH As Object
                Set WSH = CreateObject("WScript.Shell")
                WSH.Popup j & "個、ぜんぷありました", 1, "Title", vbInformation
                Set WSH = Nothing
            Else
                ' MsgBox buf & vbLf & "はありませんでした", vbApplicationModal
                msg = InputBox(Prompt:="", Default:=j & "個ありましたが、" & buf & " はありませんでした")
            End If
            f_Result.Show vbModeless
            f_Result.TextBox1 = "ない" & vbCrLf & notFound & vbCrLf & _
            "=====================" & vbCrLf & _
            "ある" & vbCrLf & found & vbCrLf & _
            "=====================" & _
            all & vbCrLf & both
            ' WorksheetFunction.Subtotal(3, Columns(selectedCol ))
        End If
    End With
End Function

