Attribute VB_Name = "TestListArray"

Option Explicit
Option Private Module
'@IgnoreModule
'@TestModule


#If twinbasic Then
    'Do nothing
#Else

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
#End If


Public Sub ListArrayTests()

#If twinbasic Then
    Debug.Print CurrentProcedureName ; vbTab, vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If
    
    
    Debug.Print "Testing completed "
    
End Sub

'@TestMethod("ListArray")
Private Sub Test01a_IsListArrayObject()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant = Array(True, True, True)
    
    Dim myResult As Variant
    ReDim myResult(0 To 2)
    Dim myArray As Variant = Array(10, 20, 30, 40, 50)
    Dim myLA As ListArray = ListArray.Deb(myArray)
    
    'Act:
    myResult(0) = VBA.IsObject(myLA)
    myResult(1) = "ListArray" = TypeName(myLA)
    myResult(2) = "ListArray" = myLA.Typename

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ListArray")
Private Sub Test01b_IsListArrayObjectByHelperToLA()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant = Array(True, True, True)
    
    Dim myResult As Variant
    ReDim myResult(0 To 2)
    Dim myArray As Variant = Array(10, 20, 30, 40, 50)
    Dim myLA As ListArray = ToLA(myArray)
    
    'Act:
    myResult(0) = VBA.IsObject(myLA)
    myResult(1) = "ListArray" = TypeName(myLA)
    myResult(2) = "ListArray" = myLA.Typename

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("ListArray")
Private Sub Test02_GetItem()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long = 30&
    
    Dim myArray As Variant = Array(10&, 20&, 30&, 40&, 50&)
    Dim myLA As ListArray = ToLA(myArray)
    Dim myResult As Variant
    
    'Act:
    myResult = myLA.Item(2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ListArray")
Private Sub Test03_LetItem()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant = Array(10&, 20&, 3000&, 40&, 50&)
    
    Dim myArray As Variant = Array(10&, 20&, 30&, 40&, 50&)
    Dim myLA As ListArray = ToLA(myArray)
    Dim myResult As Variant
    
    'Act:
    myLA.Item(2) = 3000&
    myResult = myArray
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ListArray")
Private Sub Test04_GetItemStringArray()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String = "World"
    
    Dim myArray(0 To 4)  As String
    myArray(0) = "Hello"
    myArray(1) = "There"
    myArray(2) = "World"
    myArray(3) = "Forty"
    myArray(4) = "Two"
    Dim myLA As ListArray = ToLA(CVar(myArray))
    Dim myResult As String
    
    'Act:
    myResult = myLA.Item(2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ListArray")
Private Sub Test05_LetItem()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String = "Planet"
    
    Dim myArray(0 To 4)  As String
    myArray(0) = "Hello"
    myArray(1) = "There"
    myArray(2) = "World"
    myArray(3) = "Forty"
    myArray(4) = "Two"
    Dim myLA As ListArray = ToLA(myArray)
    Dim myResult As String
    
    'Act:
    myLA.Item(2) = "Planet"
    myResult = myArray
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub