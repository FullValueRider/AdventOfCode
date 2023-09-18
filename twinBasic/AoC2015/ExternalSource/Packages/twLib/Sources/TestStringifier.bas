Attribute VB_Name = "TestStringifier"
Option Explicit
Option Private Module
'@IgnoreModule
'@TestModule
'@Folder("Tests")

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


Public Sub StringifierTests()

#If twinbasic Then

    Debug.Print CurrentProcedureName ; vbTab, vbTab,
    
#Else

    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
    
#End If



Test02a_StringifyAdmin_Empty
Test02b_StringifyAdmin_Nothing
Test02c_StringifyAdmin_Null
'Test02d_StringifyAdminUnknown  ' can't work out how to test
'Test02e_StringifyAdminError    ' can't work out how to test

Test03a_StringifyNonEnumerableObjectWithNoRecognisedMember
Test03b_StringifyNonEnumerableObjectWithDefaultMember
Test03c_StringifyNonEnumerableObjectWithTryMember

Test04a_StringifyItemByForEach_Collection
Test04b_StringifyItemByToArrayForEach_Seq
Test04c_StringifyItemByToArrayForEach_Queue
'Test04d_IterableItemByToArrayForEachStack

Test05a_StringifyKeyByForEach_Hkvp
Test05b_StringifyItemByKeysForEach_KVPair
Test05c_StringifyArray_Array_1_2_3

Test06b_ToString_Primitives1_2_3


Test07a_ArrayMarkup_NoBrackets_Array_1_2_3

Test08A_ToString_Number
Test08b_ToString_Boolean
Test08c_ToString_String

Test08f_ToString_AdminNothing

Debug.Print "Testing completed"

End Sub

'@TestMethod("Stringifier")
Public Sub Test02a_StringifyAdmin_Empty()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "Empty"
    
    Dim myResult As String
    Dim myAdmin As Variant
    myAdmin = Empty
    
    'Act:
    myResult = Stringifier.StringifyAdmin(myAdmin)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test02b_StringifyAdmin_Nothing()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "{Nothing}"
    
    
    Dim myResult As String
    Dim myAdmin As Variant
    Set myAdmin = Nothing
    
    'Act:
    myResult = Stringifier.Deb.StringifyAdmin(myAdmin)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test02c_StringifyAdmin_Null()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "Null"
    
    
    Dim myResult As String
    Dim myAdmin As Variant
    myAdmin = Null
    'Act:
    myResult = Stringifier.StringifyAdmin(myAdmin)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test02d_StringifyAdminUnknown()
    
    '************************************************************************
    'Can't currently test as I don't know how to create an object of unknown.
    '************************************************************************
    ''On Error GoTo TestFail
    
    ' 'Arrange:
    ' Dim myExpected  As String
    ' myExpected = "Unknown"
    
    
    ' Dim myResult As String
    ' Dim myAdmin As Variant
    ' Set myAdmin = New stdDataObject
    ' 'Act:
    ' myResult = Stringifier.StringifyAdmin(myAdmin, VBA.LCase$(VBA.TypeName(myAdmin)))

    ' 'Assert:
    ' AssertStrictAreEqual myExpected, myResult  , myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

' '@TestMethod("Stringifier")
' Public Sub Test02e_StringifyAdminError()
'   ' Can't test 
    ' #If twinbasic Then
    '     myProcedureName = CurrentProcedureName
    '     myComponentName = CurrentComponentName
    ' #Else
    '     myProcedureName = ErrEx.LiveCallstack.ProcedureName
    '     myComponentName = ErrEx.LiveCallstack.ModuleName
    ' #End If
    
    ' 'On Error GoTo TestFail

    ' 'Arrange:
    ' Dim myExpected  As Variant
    ' myExpected = "{""Error"",11,""Division by zero""}"
    
    ' Dim myResult As String
    ' Dim myAdmin As Variant
    ''On Error Resume Next
    ' ' Test divide by zero error
    ' myAdmin = 1 / 0
    
    ' 'Act:
    ' Err.Clear
    ''On Error GoTo TestFail
    
    ' myResult = Stringifier.StringifyAdmin(myAdmin)

    ' 'Assert:
    ' AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
' End Sub


'@TestMethod("Stringifier")
Public Sub Test03a_StringifyNonEnumerableObjectWithNoRecognisedMember()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    ' Result has a default method of Status
    Dim myExpected  As String
    myExpected = "{mpDeDup}"
    
    Dim myResult As String
    Dim myObject As mpDeDup
    Set myObject = mpDeDup.Deb
    
    'Act:
    myResult = Stringifier.StringifyNonIterableObject(myObject)
   
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test03b_StringifyNonEnumerableObjectWithDefaultMember()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail

    'Arrange:
    ' Result has a default method of Status
    Dim myExpected  As String
    myExpected = "{mpDeDup}"
    
    Dim myResult As String
    Dim myObject As mpDeDup = mpDeDup.Deb
    
    'Act:
    myResult = Stringifier.StringifyNonIterableObject(myObject)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test03c_StringifyNonEnumerableObjectWithTryMember()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail

    'Arrange:
   
    Dim myExpected  As String
    myExpected = "{[,,,]}"
    
    Dim myResult As String
    ' EntityMarkup has a ToString Method
    ' which returns Left,Separator,Right
    ' the defaults for entitymarkup are '[,]'
    ' as entitymarkup is am object the
    ' output should be {[,  ,,  ]} without the spaces
    Dim myObject As EntityMarkup = EntityMarkup.Deb

    'Act:
    myResult = Stringifier.StringifyNonIterableObject(myObject)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test04a_StringifyItemByForEach_Collection()
    
    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    'On Error GoTo TestFail

    'Arrange:
    ' Result has a default method of Status
    Dim myExpected  As String
   myExpected = Char.twLCurly & "10,20,30,40" & Char.twRCurly
    
    Dim myC As Collection = New Collection
    myC.Add 10
    myC.Add 20
    myC.Add 30
    myC.Add 40
    
    Dim myResult As String
    
    'Act:
    myResult = Stringifier.StringifyItemByForEach(myC)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test04b_StringifyItemByToArrayForEach_Seq()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As String
    myExpected = Char.twLCurly & "10,20,30,40" & Char.twRCurly
    
    Dim myS As Seq
    Set myS = Seq.Deb.AddItems(10, 20, 30, 40)

    Dim myResult As String
    
    'Act:
    myResult = Stringifier.StringifyItemByToArrayForEach(myS)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test04c_StringifyItemByToArrayForEach_Queue()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As String
   myExpected = Char.twLCurly & "10,20,30,40" & Char.twRCurly
    
    Dim myQ As Queue
    Set myQ = Queue.Deb
    myQ.Enqueue 10
    myQ.Enqueue 20
    myQ.Enqueue 30
    myQ.Enqueue 40
    
    Dim myResult As String
    
    'Act:
    myResult = Stringifier.StringifyItemByToArrayForEach(myQ)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub



'@TestMethod("Stringifier")
Public Sub Test05a_StringifyKeyByForEach_Hkvp()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail

    'Arrange:
    ' Result has a default method of Status
    Dim myExpected  As String
   myExpected = Char.twLCurly & """Ten"" 10,""Twenty"" 20,""Thirty"" 30,""Forty"" 40" & Char.twRCurly
    
    Dim myHkvp As Hkvp = Hkvp.Deb
    myHkvp.Add "Ten", 10
    myHkvp.Add "Twenty", 20
    myHkvp.Add "Thirty", 30
    myHkvp.Add "Forty", 40
    
    Dim myResult As String
    
    'Act:
    myResult = Stringifier.StringifyItembyKeysForEach(myHkvp)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test05b_StringifyItemByKeysForEach_KVPair()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    'On Error GoTo TestFail

    'Arrange:
    ' Result has a default method of Status
    Dim myExpected  As String
   myExpected = Char.twLCurly & """Ten"" 10" & Char.twRCurly
    
    Dim myIterable As KVPair = KVPair.Deb("Ten", 10)
   
    Dim myResult As String
    
    'Act:
    myResult = Stringifier.StringifyItembyKeysForEach(myIterable)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test05c_StringifyArray_Array_1_2_3()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "[1,2,3]"
    
    Dim myResult As String
    Stringifier.SetArrayMarkup
    
    'Act:
    myResult = Stringifier.StringifyArray(Array(1, 2, 3))
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test06b_ToString_Primitives1_2_3()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "1,2,3"
    
    Dim myResult As String
    
    'Act:
    myResult = Stringifier.ToString(1, 2, 3)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub



'@TestMethod("Stringifier")
Public Sub Test07a_ArrayMarkup_NoBrackets_Array_1_2_3()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "1,2,3"
    
    Dim myResult As String
   
    'Act:
    myResult = Stringifier.Deb.SetArrayMarkup(Char.twNoString, Char.twComma, Char.twNoString).StringifyArray(Array(1, 2, 3))
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub



'@TestMethod("Stringifier")
Public Sub Test08A_ToString_Number()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "42"
    
    Dim myResult As String
    myResult = Stringifier.ToString(42)
    
    'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test08b_ToString_Boolean()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "True"
    
    Dim myResult As String
    myResult = Stringifier.ToString(True)
    
    'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test08c_ToString_String()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "Hello World"
    
    Dim myResult As String
    myResult = Stringifier.ToString("Hello World")
    
    'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test08d_ToString_AdminEmpty()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "Empty"
    
    Dim myResult As String
    myResult = Stringifier.ToString(Empty)
    
    'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test08e_ToString_AdminNull()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "Null"
    Dim myResult As String
    myResult = Stringifier.ToString(Null)
    
    'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test08f_ToString_AdminNothing()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "{Object: Nothing}"
    
    Dim myNothing As Object = Nothing
    Dim myResult As String
    myResult = Stringifier.ToString(myNothing)
    
    'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test08e_ToString_Array_1_2_3()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "[1,2,3]"
    
    Dim myResult As String
    myResult = Stringifier.ToString(Array(1, 2, 3))
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Stringifier")
Public Sub Test08f_ToString_Hkvp()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String = Char.twLCurly & """Ten"" 10,""Twenty"" 20,""Thirty"" 30,""Forty"" 40" & Char.twRCurly
    
    Dim myH As Hkvp = Hkvp.Deb
    myH.Add "Ten", 10
    myH.Add "Twenty", 20
    myH.Add "Thirty", 30
    myH.Add "Forty", 40
    Dim myResult As String
    myResult = Stringifier.ToString(myH)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Stringifier")
Public Sub Test08A_ToString_ArrayOfEmpty()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "[Empty,Empty,Empty]"
    
    Stringifier.SetArrayMarkup
    Dim myResult As String
    myResult = Stringifier.StringifyArray(Array(Empty, Empty, Empty))
    
    'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub
