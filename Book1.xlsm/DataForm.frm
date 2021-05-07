VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "DataForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim obj
    Dim situArray As Object

    '室名のコンボボックス初期化
    Dim db As DbAccess
    Set db = New DbAccess
    Call db.connectDB
    Set situArray = db.getSitu
    For Each obj In situArray
        Me.situComBox.AddItem obj
    Next
    Call db.closeDB
    
End Sub

'室名コンボボックスを選択
Private Sub situComBox_click()
    '名前コンボボックスのリセット
    Me.nameComBox.Clear
End Sub

'名前コンボボックスドロップボタン押す
Private Sub nameComBox_DropButtonClick()
    Dim obj
    Dim nameArray As Object
    
    If nameComBox.ListCount = 0 Then
        '名前のコンボボックス初期化
        Dim db As DbAccess
        Set db = New DbAccess
        Call db.connectDB
        Dim test
        If Me.situComBox.ListIndex <> -1 Then
            Set nameArray = db.getNameFromSitu(Me.situComBox.List(Me.situComBox.ListIndex))
            For Each obj In nameArray
                Me.nameComBox.AddItem obj
            Next
        End If
        Call db.closeDB
    End If
End Sub

