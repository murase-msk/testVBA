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

'ADOコネクション
Public adoCn As Object
'レコードセットオブジェクト
Public adoRs As Object

' 初期化処理
Private Sub Class_Initialize()
End Sub

' メソッド
Public Sub connectDB()
    Dim strFileName As String
    strFileName = "test.accdb"
    
'    Dim adoCn As Object
    Set adoCn = CreateObject("ADODB.Connection")
    adoCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\" & strFileName & ";" 'Accessファイルに接続
    
'    Dim adoRs As Object 'ADOレコードセットオブジェクト
    Set adoRs = CreateObject("ADODB.Recordset") 'ADOレコードセットオブジェクトを作成

End Sub

'データ取り出し
Public Sub getData()
    Dim strSQL As String
    strSQL = "SELECT * FROM table1"
    adoRs.Open strSQL, adoCn 'SQLを実行して対象をRecordSetへ
    
    Do Until adoRs.EOF
        Debug.Print adoRs!フィールド1 & " " & adoRs!フィールド2
        adoRs.MoveNext
    Loop
    
    'Worksheets(1).Range("A1").CopyFromRecordset adoRs
End Sub

'
Public Sub closeDB()
    Set adoRs = Nothing
    Set adoCn = Nothing  'オブジェクトの破棄
End Sub


'''''''''''''''''''''''''''
'室名を取得する return Dictionary
Public Function getSitu() As Object
    Dim strSQL As String
    strSQL = "SELECT 室名 FROM member"
    adoRs.Open strSQL, adoCn 'SQLを実行して対象をRecordSetへ
    
    Dim situDic As Object
    Set situDic = CreateObject("Scripting.Dictionary")
    Do Until adoRs.EOF
        Dim str As String
        str = adoRs!室名
        If Not situDic.exists(str) Then
            situDic.Add str, str
        End If
        adoRs.MoveNext
    Loop
    
    Set getSitu = situDic
End Function

'室名から名前を取得する
Public Function getNameFromSitu(situ As String) As Object
    Dim strSQL As String
    strSQL = "SELECT 名前,室名 FROM member where 室名=""" & situ & """"
    adoRs.Open strSQL, adoCn 'SQLを実行して対象をRecordSetへ
    
    Dim nameDic As Object
    Set nameDic = CreateObject("Scripting.Dictionary")
    Do Until adoRs.EOF
        Dim str As String
        str = adoRs!名前
        nameDic.Add str, str
        adoRs.MoveNext
    Loop
    
    
    Set getNameFromSitu = nameDic
End Function

' 引数付き
'Public Sub SayHelloTo(otherPersonName As String)
'    MsgBox "Hello, " & otherPersonName & "!" & vbNewLine & _
'        "I'm " & Name & "!"
'End Sub

' 戻り値あり（関数）
'Public Function GetAge() As Long
'    GetAge = Age
'End Function

' オブジェクトを返す関数（Setをつけること！）
'Public Function GetMother() As Person
'    Set GetMother = Mother
'End Function

