VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "DataForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "DataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim obj
    Dim situArray As Object

    '�����̃R���{�{�b�N�X������
    Dim db As DbAccess
    Set db = New DbAccess
    Call db.connectDB
    Set situArray = db.getSitu
    For Each obj In situArray
        Me.situComBox.AddItem obj
    Next
    Call db.closeDB
    
End Sub

'�����R���{�{�b�N�X��I��
Private Sub situComBox_click()
    '���O�R���{�{�b�N�X�̃��Z�b�g
    Me.nameComBox.Clear
End Sub

'���O�R���{�{�b�N�X�h���b�v�{�^������
Private Sub nameComBox_DropButtonClick()
    Dim obj
    Dim nameArray As Object
    
    If nameComBox.ListCount = 0 Then
        '���O�̃R���{�{�b�N�X������
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

