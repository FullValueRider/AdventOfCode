Attribute VB_Name = "TestLookup"

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


Public Sub LookupTests()

#If twinbasic Then

    Debug.Print CurrentProcedureName ; vbTab, vbTab, vbTab,
    
#Else

    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
    
#End If
    
    T01_IsLookupObject
    T02_LookupAddNumberKeys
    T03_LookupAddNStringKeys
    T03_LookupKeysAndItems
    
    Debug.Print "Testing completed "
    
End Sub


'@TestMethod("Lookup")
Public Sub T01_IsLookupObject()

 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
   'On Error GoTo TestFail
    'Arrange
    Dim myExpected As String = "Lookup"
    
    'Act
    Dim myL As Lookup = Lookup.Deb
    
    'Assert
    AssertStrictAreEqual True, VBA.IsObject(myL), myProcedureName
    AssertStrictAreEqual myExpected, VBA.TypeName(myL), myProcedureName
    AssertStrictAreEqual myExpected, myL.TypeNAme, myProcedureName
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

'@TestMethod("Lookup")
Public Sub T02_LookupAddNumberKeys()

 #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    'Arrange
    Dim myExpected As Variant = Array(3, "Hello", "World", "Thing")
    
    'Act
    Dim myL As Lookup = Lookup.Deb
    myL.Add 1, "Hello"
    myL.Add 2, "World"
    myL.Add 4, "Thing"
    
    Dim myResult As Variant
    ReDim myResult(0 To 3)
    myResult(0) = myL.Count
    myResult(1) = myL.Item(1)
    myResult(2) = myL.Item(2)
    myResult(3) = myL.Item(4)
    
    'Assert
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

Public Sub T03_LookupAddNStringKeys()

 #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    'Arrange
    Dim myExpected As Variant = Array(3, 1, 2, 4)
    
    'Act
    Dim myL As Lookup = Lookup.Deb
    myL.Add "Hello", 1
    myL.Add "World", 2
    myL.Add "Thing", 4
    
    Dim myResult As Variant
    ReDim myResult(0 To 3)
    myResult(0) = myL.Count
    myResult(1) = myL.Item("Hello")
    myResult(2) = myL.Item("World")
    myResult(3) = myL.Item("Thing")
    
    'Assert
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
   
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

Public Sub T03_LookupKeysAndItems()

 #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    'Arrange
    Dim myExpectedItems As Variant = Array(1, 2, 4)
    Dim myExpectedKeys As Variant = Array("Hello", "World", "Thing")
    'Act
    Dim myL As Lookup = Lookup.Deb
    myL.Add "Hello", 1
    myL.Add "World", 2
    myL.Add "Thing", 4
    
    Dim myResultKeys As Variant = myL.Keys
    Dim myResultItems As Variant = myL.Items
    
    
    'Assert
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

Public Sub T04_LookupAddPairs()

 #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
   'On Error GoTo TestFail
    'Arrange
    Dim myExpectedItems As Variant = Array(1, 2, 4)
    Dim myExpectedKeys As Variant = Array("Hello", "World", "Thing")
    'Act
    Dim myL As Lookup = Lookup.Deb.AddPairs(Array("Hello", "World", "Thing"), Array(1, 2, 4))
   
    Dim myResultKeys As Variant = myL.Keys
    Dim myResultItems As Variant = myL.Items
    
    
    'Assert
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub