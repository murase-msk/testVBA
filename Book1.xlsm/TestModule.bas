Attribute VB_Name = "TestModule"
Sub testSub()
    
    Dim db As DbAccess
    Set db = New DbAccess
    Call db.connectDB
    Call db.getData
    Call db.closeDB
    
End Sub

Sub testArrayList()
    Dim arraylist As Object
    Dim myObj
   
    Set arraylist = CreateObject("System.Collections.ArrayList")
    arraylist.Add ("Z")
    arraylist.Add ("Y")
    arraylist.Add ("Y")
    arraylist.Add ("W")

    'arraylist.Sort

    For Each myObj In arraylist
        'MsgBox myObj
    Next
    
    Debug.Print arraylist(1)
    
    'dictionaryÇ≈èdï°çÌèú
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    For Each myObj In arraylist
        If Not dic.exists(myObj) Then
            dic.Add myObj, myObj
        End If
    Next
    
    MsgBox "OK"
        
End Sub

