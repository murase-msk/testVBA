VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ADO�R�l�N�V����
Public adoCn As Object
'���R�[�h�Z�b�g�I�u�W�F�N�g
Public adoRs As Object

' ����������
Private Sub Class_Initialize()
End Sub

' ���\�b�h
Public Sub connectDB()
    Dim strFileName As String
    strFileName = "test.accdb"
    
'    Dim adoCn As Object
    Set adoCn = CreateObject("ADODB.Connection")
    adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\" & strFileName & ";" 'Access�t�@�C���ɐڑ�
    
'    Dim adoRs As Object 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g
    Set adoRs = CreateObject("ADODB.Recordset") 'ADO���R�[�h�Z�b�g�I�u�W�F�N�g���쐬

End Sub

'�f�[�^���o��
Public Sub getData()
    Dim strSQL As String
    strSQL = "SELECT * FROM table1"
    adoRs.Open strSQL, adoCn 'SQL�����s���đΏۂ�RecordSet��
    
    Do Until adoRs.EOF
        Debug.Print adoRs!�t�B�[���h1 & " " & adoRs!�t�B�[���h2
        adoRs.MoveNext
    Loop
    
    'Worksheets(1).Range("A1").CopyFromRecordset adoRs
End Sub

'
Public Sub closeDB()
    Set adoRs = Nothing
    Set adoCn = Nothing  '�I�u�W�F�N�g�̔j��
End Sub


'''''''''''''''''''''''''''
'�������擾���� return Dictionary
Public Function getSitu() As Object
    Dim strSQL As String
    strSQL = "SELECT ���� FROM member"
    adoRs.Open strSQL, adoCn 'SQL�����s���đΏۂ�RecordSet��
    
    Dim situDic As Object
    Set situDic = CreateObject("Scripting.Dictionary")
    Do Until adoRs.EOF
        Dim str As String
        str = adoRs!����
        If Not situDic.exists(str) Then
            situDic.Add str, str
        End If
        adoRs.MoveNext
    Loop
    
    Set getSitu = situDic
End Function

'�������疼�O���擾����
Public Function getNameFromSitu(situ As String) As Object
    Dim strSQL As String
    strSQL = "SELECT ���O,���� FROM member where ����=""" & situ & """"
    adoRs.Open strSQL, adoCn 'SQL�����s���đΏۂ�RecordSet��
    
    Dim nameDic As Object
    Set nameDic = CreateObject("Scripting.Dictionary")
    Do Until adoRs.EOF
        Dim str As String
        str = adoRs!���O
        nameDic.Add str, str
        adoRs.MoveNext
    Loop
    
    
    Set getNameFromSitu = nameDic
End Function

' �����t��
'Public Sub SayHelloTo(otherPersonName As String)
'    MsgBox "Hello, " & otherPersonName & "!" & vbNewLine & _
'        "I'm " & Name & "!"
'End Sub

' �߂�l����i�֐��j
'Public Function GetAge() As Long
'    GetAge = Age
'End Function

' �I�u�W�F�N�g��Ԃ��֐��iSet�����邱�ƁI�j
'Public Function GetMother() As Person
'    Set GetMother = Mother
'End Function
