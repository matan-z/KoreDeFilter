Attribute VB_Name = "m_FilteringByArray"
Option Explicit

Public i As Long, j As Long, Col As Long, �ŏI�s As Long, selectedCol As Long, AFS As Long, AFE As Long
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
                MsgBox "�I�[�g�t�B���^���ݒ肳��Ă��܂���"
            End If
        Else
            Dim tableName As String
            tableName = ws.ListObjects(1).Name
            If ws.ListObjects(tableName).ShowAutoFilter Then
                Set myRng = ws.ListObjects(tableName).AutoFilter.Range
                Call c.FilteringByRange(myRng, f_FilterTool.TextBox1.Value)  '2-1
            Else
                MsgBox "�I�[�g�t�B���^���ݒ肳��Ă��܂���"
            End If
        End If
    Else
        MsgBox "�e�L�X�g�{�b�N�X�ɉ��������Ă��܂���"
    End If
End Sub

'�y�Q�lURL�zhttps://desmondoshiwambo.wordpress.com/2012/02/23/how-to-copy-and-paste-texttofrom-clipboard-using-vba-microsoft-access/
Sub CopyTextToClipboard(ByVal inText As String)
    Dim objClipboard As Object
    Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    objClipboard.SetText inText
    objClipboard.PutInClipboard
    Set objClipboard = Nothing
End Sub

Sub �e�L�X�g�{�b�N�X�ŃJ���}��؂�(ByVal text As String)
    Dim CB As Variant
    Dim �N���b�v() As String
    �N���b�v = ���s�Ŋi�[(text)
    Dim �N���b�v�z��() As String, i As Long
    ReDim �N���b�v�z��(UBound(�N���b�v))
    For i = 0 To UBound(�N���b�v)
        If �N���b�v(i) <> "" Then
            �N���b�v�z��(i) = Replace(�N���b�v(i), vbCr, "")
        End If
    Next i
    '�y�Q�lURL�zhttps://oshiete.goo.ne.jp/qa/6871211.html
    Dim c As Variant
    Dim �� As Long
    Dim �z�� As Variant
    Set DB = CreateObject("Scripting.Dictionary")
    For Each c In �N���b�v�z��
        DB(c) = 1
    Next c
    �� = DB.Count '�Q�l�A�V�z��̌�
    �z�� = DB.Keys '�V�z��̓W�J
    buf = ""
    For i = 0 To UBound(�z��)
        buf = buf + �z��(i) & ","
    Next i
    If Len(buf) > 0 Then
        buf = Left(buf, Len(buf) - 1)
    End If
    CopyTextToClipboard (buf)
    MsgBox buf & vbLf & "���N���b�v�{�[�h�ɒu���܂���"
End Sub
Sub �󔒍폜(ByVal text As String)
    Dim CB As Variant
    Dim �N���b�v() As String
    �N���b�v = ���s�Ŋi�[(text)
    Dim �N���b�v�z��() As String, i As Long
    ReDim �N���b�v�z��(UBound(�N���b�v))
    For i = 0 To UBound(�N���b�v)
        �N���b�v�z��(i) = Replace(�N���b�v(i), vbCr, "")
    Next i
    '�d���폜
    '�y�Q�lURL�zhttps://oshiete.goo.ne.jp/qa/6871211.html
    Dim c As Variant
    Dim �� As Long
    Dim �z�� As Variant
    Set DB = CreateObject("Scripting.Dictionary")
    For Each c In �N���b�v�z��
        DB(c) = 1
    Next c
    �� = DB.Count
    �z�� = DB.Keys
    i = 0
    j = 0
    buf = ""
    For i = 0 To UBound(�z��)
        If �z��(i) <> "" And i <> UBound(�z��) Then
            buf = buf & �z��(i) & vbLf
            j = j + 1
        Else
            If �z��(i) <> "" And i = UBound(�z��) Then
                buf = buf & �z��(i)
                j = j + 1
            End If
        End If
    Next
    f_FilterTool.TextBox1.Value = ""
    f_FilterTool.TextBox1.Value = buf
End Sub

Function ���s�Ŋi�[(str As String) As String()
    ���s�Ŋi�[ = Split(TrimLF(str), vbLf)
End Function
Sub �����Z���I���̏ꍇ()
    Dim ad, buf As String
    ad = Selection.Address
    Dim adArray() As String
    adArray = Split(ad, ",")
    Dim �z�� As Variant
    ReDim �z��(UBound(adArray))
    buf = ""
    For i = 0 To UBound(adArray)
        �z��(i) = Range(adArray(i)).Value
        If i <> UBound(adArray) Then
            buf = buf + Range(adArray(i)).Value '+ ","
        Else
            buf = buf + Range(adArray(i)).Value
        End If
    Next i
    With New MSForms.DataObject
        .SetText buf '�ϐ��̒l��DataObject�Ɋi�[����
        .PutInClipboard 'DataObject�̃f�[�^���N���b�v�{�[�h�Ɋi�[����
    End With
End Sub
'�y�Q�lURL�zhttps://desmondoshiwambo.wordpress.com/2012/02/23/how-to-copy-and-paste-texttofrom-clipboard-using-vba-microsoft-access/

'(does not require reference) :
Function GetTextFromClipboard() As String
    Dim objClipboard As Object
    Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    objClipboard.GetFromClipboard
    On Error GoTo myError
    GetTextFromClipboard = objClipboard.GetText
    '******************************�G���[����
myError:
    If Err.Number <> 0 Then '2017/06/11�ɒǋL �G���[�[���ŗ\�����ʃG���[�ƂȂ邽��
        Select Case Err.Number
            Case -2147221404
                MsgBox "�N���b�v�{�[�h����ł�" & vbCrLf & Err.Description, vbExclamation
                '�N���b�v�{�[�h�Ńt�B���^.Show vbModeless
            End
            Case Else
                On Error Resume Next
                '�N���b�v�{�[�h�Ńt�B���^.Show vbModeless
                MsgBox "�\�����ʃG���[���������܂���!", vbExclamation
        End Select
    End If
    '******************************
    Set objClipboard = Nothing
End Function
