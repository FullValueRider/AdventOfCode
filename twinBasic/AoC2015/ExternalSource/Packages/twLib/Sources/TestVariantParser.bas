Attribute VB_Name = "TestVariantParser"
Option Explicit

#If twinbasic Then
    'Do nothing
#Else
'@IgnoreModule
'@TestModule
Option Private Module

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


Public Sub VariantParserTests()

#If twinbasic Then
    Debug.Print CurrentProcedureName ; vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If

    Test01_NewVariantParser
    Test02_ToForEachSinglePrimitive
    Test03_ToForEachSingleNonEnumerableObject
    Test04a_ToForEach_EmulatedParamrray
    Test04b_ToForEachVariantEncapsulatedArray
    Test05_ToForEachCollection
    Test06_ToForEachSeq
    Test07_ToForEachHkvp
    Test08_ToForEachPreservedString
    Test09_ToForEachStringToArray
    Test10_ToForEachStringToCharSeq
    Test11_ToForEachStringToCharArray
    'The null cases return Array(Empty)
    Test12_ToForEachOmitDeb
    Test13_ToForEachSingleEmpty
    Test14_ToForEachEnumerableWithCountZero
    Debug.Print "Testing completed"

End Sub

'@TestMethod("VariantParser")
Private Sub Test01_NewVariantParser()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(True, "VariantParser", "VariantParser")

    '@Ignore IntegerDataType
    Dim myVP As VariantParser = VariantParser.Deb(10)
   
    
    Dim myResult As Variant
    ReDim myResult(0 To 2)
    
    'Act:
    myResult(0) = VBA.IsObject(myVP)
    myResult(1) = VBA.TypeName(myVP)
    myResult(2) = myVP.Name
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test02_ToForEachSinglePrimitive()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    
    ' Cargo                               As Variant
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 
    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 1, OfNumbers, OfArray, "long", idLong)

    Dim myExpectedCargo As Variant
    myExpectedCargo = Array(10&)
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
   
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb(10&).ToForEach(StringAsString)
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictSequenceEquals myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test03_ToForEachSingleNonEnumerableObject()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    
    'Arrange:
    
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 
    
    ' mpInc is not a registered type so sHashd will return null
    ' when the name "mpinc" is not found, hence idEmpty (0)
    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 1, OfItemObjects, OfArray, "mpinc", idEmpty)

    Dim myExpectedCargo As Variant
    myExpectedCargo = Array(11)
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    Dim mympInc As mpInc = mpInc.Deb
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb(mympInc).ToForEach
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = Array(myPR.Cargo(0).execmapper(10))
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictSequenceEquals myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test04a_ToForEach_EmulatedParamrray()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 
    
    ' p.Cargo = Empty
    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 9, ofParamArray, OfArray, "variant", idVariant)

    Dim myExpectedCargo As Variant
    myExpectedCargo = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    Dim myResultCargo As Variant
   
    'Act:
    ' The array is emulating a forwarded ParamArray so IsArray should be false
    Dim myPR As ParserResult = VariantParser.Deb(Array(10, 20, 30, 40, 50, 60, 70, 80, 90)).ToForEach
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictSequenceEquals myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
'@TestMethod("VariantParser")
Private Sub Test04b_ToForEachVariantEncapsulatedArray()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 

    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 9, OfArray, OfArray, "variant", idVariant)

    Dim myExpectedCargo As Variant
    myExpectedCargo = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    ' The array is emulating a forwarded ParamArray containing a single item which is an array
    Dim myPR As ParserResult = VariantParser.Deb(Array(Array(10, 20, 30, 40, 50, 60, 70, 80, 90))).ToForEach
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictSequenceEquals myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VariantParser")
Private Sub Test05_ToForEachCollection()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 
    
    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 9, OfItemByForEach, OfItemByForEach, "collection", idCollection)
    Dim myC As Collection = New Collection
    With myC
    
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myExpectedCargo As Variant
    Set myExpectedCargo = myC
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb(myC).ToForEach
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    Set myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    Dim myIndex As Long
    For myIndex = 1 To 9
        AssertStrictAreEqual myExpectedCargo(myIndex), myResultCargo(myIndex), myProcedureName
    Next
    

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test06_ToForEachSeq()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 

    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 9, OfItemByToArrayForEach, OfArray, "seq", idSeq)
    Dim myS As Seq = Seq.Deb(10, 20, 30, 40, 50, 60, 70, 80, 90)
    
    Dim myExpectedCargo As Variant
    myExpectedCargo = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb(myS).ToForEach
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictSequenceEquals myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test07_ToForEachHkvp()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 

    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 9, OfItemByKeysForeach, OfArray, "hkvp", idHkvp)
    Dim myH As Hkvp = Hkvp.Deb.AddPairs(VBA.Split("a,b,c,d,e,f,g,h,i", ","), Array(10, 20, 30, 40, 50, 60, 70, 80, 90))
    
    Dim myExpectedCargo As Variant
    myExpectedCargo = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb(myH).ToForEach
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictSequenceEquals myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test08_ToForEachPreservedString()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 

    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 1, OfStrings, OfStrings, "string", idString)
    
    Dim myExpectedCargo As Variant
    myExpectedCargo = "Hello"
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb("Hello").ToForEach(StringAsString)
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictAreEqual myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test09_ToForEachStringToArray()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    
    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 

    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 1, OfStrings, OfArray, "string", idString)
    
    
    Dim myExpectedCargo As Variant
    myExpectedCargo = "Hello"
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb("Hello").ToForEach(StringToArray)
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo(0)
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictAreEqual myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test10_ToForEachStringToCharSeq()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    
    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 

    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 5, OfStrings, OfItemByToArrayForEach, "string", idString)
    
    
    Dim myExpectedCargo As Variant
    myExpectedCargo = Array("H", "e", "l", "l", "o")
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb("Hello").ToForEach(StringToCharSeq)
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo.toarray
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictSequenceEquals myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test11_ToForEachStringToCharArray()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    
    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 

    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 5, OfStrings, OfArray, "string", idString)
    
    
    Dim myExpectedCargo As Variant
    myExpectedCargo = Array("H", "e", "l", "l", "o")
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb("Hello").ToForEach(StringToCharArray)
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictSequenceEquals myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test12_ToForEachOmitDeb()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    
    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 

    ' not calling deb will result an a cargo of admin empty
    Dim myExpectedData  As Variant
    myExpectedData = Array(False, -1, OfAdmins, ofNoGroup, "empty", idEmpty)
    
    
    Dim myExpectedCargo As Variant
    myExpectedCargo = Empty
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.ToForEach(StringToCharArray)
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictAreEqual myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test13_ToForEachSingleEmpty()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

       
    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 

    Dim myExpectedData  As Variant
    myExpectedData = Array(False, -1, OfAdmins, ofNoGroup, "empty", idEmpty)
    
    
    Dim myExpectedCargo As Variant
    myExpectedCargo = Empty
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb(Empty).ToForEach(StringToCharArray)
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictAreEqual myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VariantParser")
Private Sub Test14_ToForEachEnumerableWithCountZero()
    
#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

       
    'On Error GoTo TestFail
    'Arrange:
    'IsQueryable As boolean ' True of the variant contained valid data
    'Count
    'InputGroup 
    'ResultGroup 
    'InputBaseType
    'InputBaseOrd 

    Dim myExpectedData  As Variant
    myExpectedData = Array(True, 0, OfItemByToArrayForEach, OfArray, "seq", idSeq)
    
    Dim mySeq As Seq = Seq.Deb
    
    Dim myExpectedCargo As Variant
    myExpectedCargo = Array(Empty)
    
    Dim myResultData As Variant
    ReDim myResultData(0 To 5)
    
    Dim myResultCargo As Variant
   
    'Act:
    Dim myPR As ParserResult = VariantParser.Deb(mySeq).ToForEach(StringToCharArray)
    
    myResultData(0) = myPR.IsAllocated
    myResultData(1) = myPR.Count
    myResultData(2) = myPR.InputGroup
    myResultData(3) = myPR.ResultGroup
    myResultData(4) = myPR.InputBaseType
    myResultData(5) = myPR.InputBaseOrd
    
    myResultCargo = myPR.Cargo
    
    'Assert:
    AssertStrictSequenceEquals myExpectedData, myResultData, myProcedureName
    AssertStrictSequenceEquals myExpectedCargo, myResultCargo, myProcedureName

TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'ToDo: Add tests for method 'ToItems'