VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_Result 
   Caption         =   "Result"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   OleObjectBlob   =   "f_Result.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "f_Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Call m_FilteringByArray.CopyTextToClipboard(f_Result.TextBox1.Value)
End Sub
Private Sub CommandButton2_Click()
    Dim Path As String, WSH As Variant
    Set WSH = CreateObject("WScript.Shell")
    Path = WSH.SpecialFolders("Desktop") & "\"
    Set WSH = Nothing
    Dim TxtFullPath As String
    TxtFullPath = Path & "フィルター結果_" & Replace(Date, "/", "-") & "_" & Replace(Time, ":", "-") & ".txt"
    Open TxtFullPath For Output As #1
    'UserForm3.TextBox12の値をファイルに書き込む
    Write #1, f_Result.TextBox1.Value
    '更新情報.txtを閉じる
    Close #1
    Dim fso As Object
    Dim sfile As String
    Set fso = CreateObject("shell.application")
    sfile = TxtFullPath
    fso.Open (sfile)
End Sub
Private Sub UserForm_Click()
End Sub
Private Sub UserForm_Initialize()
    With TextBox1
    '      .text = "複数行入力サンプル　-jizilog.com-　"
          .MultiLine = True
    End With
End Sub
