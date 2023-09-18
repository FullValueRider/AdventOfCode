Attribute VB_Name = "TestTypeInfo"
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

    

Public Sub TypeInfoTests()
    
#If twinbasic Then
    Debug.Print CurrentProcedureName ; vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If

    Test01a_IsNumber
    Test01b_IsNotNumber
    
    Test02a_IsContainer
    Test02b_IsNotContainer
 
    Test03a_IsItemObject
    Test03b_IsNotItemObject
    
    Test04a_IsString
    Test04b_IsNotString
    
    Test05a_IsAdmin
    Test05b_IsNotAdmin
    
    Test06a_IsPrimitive
    Test06b_IsNotPrimitive
    
    Test07a_IsItemByIndex0
    Test07b_IsNotItemByIndex0
    
    Test08a_IsItemByIndex1
    Test08b_IsNotItemByIndex1
    
    Test09a_IsItemByKey
    Test09b_IsNotItemByKey
    
    Test10a_IsItemByArray
    Test10b_IsNotItemByArray
    
    
    
    
    
    
    
'     Test03a_IsTypeByTypeId_MultipleIntegers_True
'     Test03b_IsTypeByTypeId_MultipleIntegers_OneSingle_False
'     Test03c_IsTypeByTypeName_MultipleIntegers_True
'     Test03d_IsTypeByTypeName_MultipleIntegers_OneSingle_False
'     Test03e_IsTypeByIntegerType_MultipleIntegers_True
'     Test03f_IsTypeByHkvp_MultipleIntegers_String_False
    
'    ' Test04_MultipleIntegersWithSingleNonIntegerIsFalse
'     ' Test05_MultipleIntegersWithIntegerArrayIsFalse
'     ' Test06_MultipleIntegersWithBoxedIntegerArrayIsFalse
'     Test07_MultipleIntegersWithVariantArrayofIntegerIsFalse
'     'Test08_MultipleIntegersWithVariantArrayWithVariantArraysOfInteger
    
'     Test11_SingleNonIntegerIsFalse
'     Test12_MultipleIntegersIsFalse
'     Test13_MultipleNonIntegersIsTrue
'     Test14_MultipleIntegersWithSingleNonIntegerIsFalse
'     Test15_MultipleNonIntegersWithNonIntegerArrayIsTrue
'     ' Test16_MultipleNonIntegersWithBoxedNonIntegerArrayIsTrue
'     Test17_MultipleNonIntegersWithVariantArrayofNonIntegerIsTrue
'     Test18_MultipleIntegersWithVariantArrayWithVariantArraysOfInteger
    
'     Test21_SingleArrayIsTrue
'     Test22_MultipleArraysIsTrue
'     Test23_IntegerIsFalse
'     Test24_IntegerAndMultipleArraysIsFalse
    
'     Test31_SingleIntgerIsTrue
'     Test32_MultipleNonArraysIsTrue
'     Test33_VariantArrayIsFalse
'     Test34_IntegerAndMultipleArraysIsFalse
    
'     Test41_SingleObjectIsTrue
'     Test42_MultipleObjectsIsTrue
'     Test43_IntegerIsFalse
'     Test44_VariantIsFalse
'     Test45_IntegerWithMultipleObjectsIsFalse
    
'     Test51_SingleIntgerIsTrue
'     Test52_MultipleNonObjectssIsTrue
'     Test53_VariantisTrue
'     Test54_IntegerAndMultipleObjectsIsTrue
    
'     Test61_IntegerArrayIsTrue
'     Test62_VariantIsFalse
'     Test63_VariantWithIntegerArrayIsTrue
'     Test64_VariantArrayNotUniformIsFalse
'     Test65_UniformVariantArrayOfVariantArrayIsTrue
'     Test66_NonUniformVariantArrayOfVariantArrayIsFalse
    
'     Test71_IntegerArrayOfLongIsFalse
'     Test72_VariantOfLongIsFalse
'     Test73_LongVariantWithIntegerArrayIsTrue
'     Test74_VariantArrayNotUniformIsTrue
'     Test75_UniformVariantArrayOfVariantArrayIsTrue
'     Test76_NonUniformVariantArrayOfVariantArrayIsTrue
'     ' Test01_HasItemsEmptyArrayOfIntegerIsFalse
'     ' Test02_HasItemsArrayOfIntegerIsTrue
'     ' Test03_HoldsItemVariantHoldingEmptyIsFalse
'     ' Test04_HoldsItemVariantIsNullArrayIsFalse
'     ' Test05_HoldsItemVariantIsArrayOfIntegerIsTrue
'     ' Test06_HoldsItemArrayListIsNothingIsFalse
'     ' Test07_HoldsItemArrayListIsPopulatedTrue
'     ' Test08_HoldsItemArrayListIsNothingIsFalse
'     ' Test09_HoldsItemArrayListIsPopulatedTrue

' '    Test10_CountEmptyArray
' '    Test11_CountArray
' '    Test12_CountEmptyArrayList
' '    Test13_CountArrayList
' '    Test14_CountEmptyCollection
' '    Test15_CountCollection
' '    Test16_CountEmptyQueue
' '    Test17_CountQueue
' '    Test18_CountEmptyStack
' '    Test19_CountStack
' '    Test20_CountEmptyArray
' '    Test21_CountArray
' '    Test22_CountEmptyArrayList
' '    Test23_CountArrayList
' '    Test24_CountEmptyCollection
' '    Test27_CountCollection
' '    Test28_CountEmptyQueue
' '    Test29_CountQueue
' '    Test30_CountEmptyStack

'     ' Test31_TryExtentWithEmptyArray
'     ' Test32_TryExtentWithFilledArray
'     ' Test33_TryExtentWithEmptyArrayList
'     ' Test34_TryExtentWithFilledArrayList
'     ' Test35_TryExtentWithEmptyCollection
'     ' Test36_TryExtentWithFilledCollection
'     ' Test37_TryExtentWithEmptyQueue
'     ' Test38_TryExtentWithFilledQueue
'     ' Test39_TryExtentWithEmptyStack
'     ' Test40_TryExtentWithFilledStack
    

' '    Test51_LastIndexTryGetFromEmptyArray
' '    Test52_LastIndexTryGetFromArray
' '    Test53_LastIndexTryGetFromEmptyArrayList
' '    Test54_LastIndexTryGetFromArrayList
' '    Test55_LastIndexTryGetFromEmptyCollection
' '    Test56_LastIndexTryGetFromCollection
' '    Test57_LastIndexTryGetFromEmptyQueue
' '    Test58_LastIndexTryGetFromQueue
' '    Test59_LastIndexTryGetFromEmptyStack
' '    Test60_LastIndexTryGetFromStack
' '    Test61_LastIndexGetFromEmptyArray
' '    Test62_LastIndexGetFromArray
' '    Test63_GetFromEmptyArrayList
' '    Test64_LastIndexGetFromArrayList
' '    Test65_LastIndexGetFromEmptyCollection
' '    Test66_LastIndexGetFromCollection
' '    Test67_LastIndexGetFromEmptyQueue
' '    Test68_LastIndexGetFromQueue
' '    Test69_LastIndexGetFromEmptyStack
' '    Test70_LastIndexGetFromStack

'     ' Test71_ToArrayFromArray
'     ' Test72_ToArrayFromArrayList
'     ' Test73_ToArrayFromCollection
'     ' Test74_ToArrayFromQueue
'     ' Test75_ToArrayFromStack

'     ' Test76_ToArrayListFromArray
'     ' Test77_ToArrayListFromArrayList
'     ' Test78_ToArrayListFromCollection
'     ' Test79_ToArrayListFromQueue
'     ' Test80_ToArrayListFromStack

'     ' Test81_ToCollectionFromArray
'     ' Test82_ToCollectionFromArrayList
'     ' Test83_ToCollectionFromCollection
'     ' Test84_ToCollectionFromQueue
'     ' Test85_ToCollectionFromStack

'     ' Test86_ToListFromArray
'     ' Test87_ToListFromArrayList
'     ' Test88_ToListFromList
'     ' Test89_ToListFromQueue
'     ' Test90_ToListFromStack

'     ' Test91_ToQueueFromArray
'     ' Test92_ToQueueFromArrayList
'     ' Test93_ToQueueFromCollection
'     ' Test94_ToQueueFromQueue
'     ' Test95_ToQueueFromStack

'     ' Test96_ToStackFromArray
'     ' Test97_ToStackFromArrayList
'     ' Test98_ToStackFromCollection
'     ' Test99_ToStackFromQueue
'     ' Test100_ToStackFromStack

'     ' Test101_MinMaxFromArray
'     ' Test102_MinMaxFromArrayList
'     ' Test103_MinMaxFromCollection
'     ' Test104_MinMaxFromQueue
'     ' Test105_MinMaxFromStack

'     ' Test106_SumFromArray'@TestMethod("")
'     ' Test107_SumFromArrayList
'     ' Test108_SumFromCollection
'     ' Test109_SumFromQueue
'     ' Test110_SumFromStack

    Debug.Print "Testing completed"

End Sub
    

' ''#Region "HoldsItem"



'@TestMethod("TypeInfo")
Private Sub Test01a_IsNumber()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected()  As Variant
    myExpected = Array(True, True, True, True, True, True, True, True, True, True, False)

    Dim myTest(0 To 10) As Variant
    myTest(0) = CByte(42.1)
    myTest(1) = CCur(42.1)
    myTest(2) = CDate(42.1)
    myTest(3) = CDec(42.1)
    myTest(4) = CDbl(42.1)
    myTest(5) = CInt(42.1)
    myTest(6) = CLng(42.1)
    myTest(7) = CLngLng(42.1)
    myTest(8) = CLngPtr(42.1)
    myTest(9) = CSng(42.1)
    myTest(10) = "HelloWorld"

    Dim myResult(0 To 10)  As Variant

    'Act:
    Dim myIndex As Long
    For myIndex = 0 To 10
    myResult(myIndex) = TypeInfo.IsNumber(myTest(myIndex))
    Next
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test01b_IsNotNumber()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected()  As Variant
    myExpected = Array(False, False, False, False, False, False, False, False, False, False, True)

    Dim myTest(0 To 10) As Variant
    myTest(0) = CByte(42.1)
    myTest(1) = CCur(42.1)
    myTest(2) = CDate(42.1)
    myTest(3) = CDec(42.1)
    myTest(4) = CDbl(42.1)
    myTest(5) = CInt(42.1)
    myTest(6) = CLng(42.1)
    myTest(7) = CLngLng(42.1)
    myTest(8) = CLngPtr(42.1)
    myTest(9) = CSng(42.1)
    myTest(10) = "HelloWorld"

    Dim myResult(0 To 10)  As Boolean

    'Act:
    Dim myIndex As Long
    For myIndex = 0 To 10
    myResult(myIndex) = TypeInfo.IsNotNumber(myTest(myIndex))
    Next
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TypeInfo")
Private Sub Test02a_IsContainer()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, False, False, True)

    Dim myS As Seq
    Set myS = Seq.Deb

    Dim myResult(0 To 3)  As Boolean

    'Act:
    'Test two object so we don't have to encapsulate in array
    myResult(0) = TypeInfo.IsContainer(myS)
    myResult(1) = TypeInfo.IsContainer(42)
    myResult(2) = TypeInfo.IsContainer("Hello")
    myResult(3) = TypeInfo.IsContainer(Hkvp.Deb)

    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TypeInfo")
Private Sub Test02b_IsNotContainer()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, True, True, False)

    Dim myS As Seq
    Set myS = Seq.Deb

    Dim myResult(0 To 3)  As Boolean

    'Act:
    'Test two object so we don't have to encapsulate in array
    myResult(0) = TypeInfo.IsNotContainer(myS)
    myResult(1) = TypeInfo.IsNotContainer(42)
    myResult(2) = TypeInfo.IsNotContainer("Hello")
    myResult(3) = TypeInfo.IsNotContainer(Hkvp.Deb)

    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test03a_IsItemObject()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, False, False, False)

    Dim myMapper As mpInc
    Set myMapper = mpInc.Deb(1)

    Dim myResult(0 To 3)  As Boolean

    'Act:
    'Test two object so we don't have to encapsulate in array
    myResult(0) = TypeInfo.IsItemObject(myMapper)
    myResult(1) = TypeInfo.IsItemObject(42)
    myResult(2) = TypeInfo.IsItemObject("Hello")
    myResult(3) = TypeInfo.IsItemObject(Hkvp.Deb)

    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test03b_IsNotItemObject()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, True, True, True)

    Dim myMapper As mpInc
    Set myMapper = mpInc.Deb(1)

    Dim myResult(0 To 3)  As Boolean

    'Act:
    'Test two object so we don't have to encapsulate in array
    myResult(0) = TypeInfo.IsNotItemObject(myMapper)
    myResult(1) = TypeInfo.IsNotItemObject(42)
    myResult(2) = TypeInfo.IsNotItemObject("Hello")
    myResult(3) = TypeInfo.IsNotItemObject(Hkvp.Deb)

    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TypeInfo")
Private Sub Test04a_IsString()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected()  As Variant
    myExpected = Array(False, False, False, False, False, False, False, False, False, False, True)

    Dim myTest(0 To 10) As Variant
    myTest(0) = CByte(42.1)
    myTest(1) = CCur(42.1)
    myTest(2) = CDate(42.1)
    myTest(3) = CDec(42.1)
    myTest(4) = CDbl(42.1)
    myTest(5) = CInt(42.1)
    myTest(6) = CLng(42.1)
    myTest(7) = CLngLng(42.1)
    myTest(8) = CLngPtr(42.1)
    myTest(9) = CSng(42.1)
    myTest(10) = "HelloWorld"


    Dim myResult(0 To 10)  As Boolean

    'Act:
    Dim myIndex As Long
    For myIndex = 0 To 10
    myResult(myIndex) = TypeInfo.IsString(myTest(myIndex))
    Next
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test04b_IsNotString()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

    

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected()  As Variant
    myExpected = Array(True, True, True, True, True, True, True, True, True, True, False)

    Dim myTest(0 To 10) As Variant
    myTest(0) = CByte(42.1)
    myTest(1) = CCur(42.1)
    myTest(2) = CDate(42.1)
    myTest(3) = CDec(42.1)
    myTest(4) = CDbl(42.1)
    myTest(5) = CInt(42.1)
    myTest(6) = CLng(42.1)
    myTest(7) = CLngLng(42.1)
    myTest(8) = CLngPtr(42.1)
    myTest(9) = CSng(42.1)
    myTest(10) = "HelloWorld"

    Dim myResult(0 To 10)  As Boolean

    'Act:
    Dim myIndex As Long
    For myIndex = 0 To 10
    myResult(myIndex) = TypeInfo.IsNotString(myTest(myIndex))
    Next
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TypeInfo")
Private Sub Test05a_IsAdmin()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, True, True, True, True, False)


    Dim myResult(0 To 5)  As Boolean

    'Act:
    'Test two object so we don't have to encapsulate in array
    myResult(0) = TypeInfo.IsAdmin(42)
    myResult(1) = TypeInfo.IsAdmin(Empty)
    myResult(2) = TypeInfo.IsAdmin(Nothing)
    myResult(3) = TypeInfo.IsAdmin(Null)
    myResult(4) = TypeInfo.IsAdmin(CVErr(79))
    myResult(5) = TypeInfo.IsAdmin(Seq.Deb)

    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test05b_IsNotAdmin()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, False, False, False, False, True)


    Dim myResult(0 To 5)  As Boolean

    'Act:
    'Test two object so we don't have to encapsulate in array
    myResult(0) = TypeInfo.IsNotAdmin(42)
    myResult(1) = TypeInfo.IsNotAdmin(Empty)
    myResult(2) = TypeInfo.IsNotAdmin(Nothing)
    myResult(3) = TypeInfo.IsNotAdmin(Null)
    myResult(4) = TypeInfo.IsNotAdmin(CVErr(79))
    myResult(5) = TypeInfo.IsNotAdmin(Seq.Deb)

    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test06a_IsPrimitive()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected()  As Variant
    myExpected = Array(False, True, True, True, True, True, True, True, True, True, True, True, False)

    Dim myTest(-1 To 11) As Variant
    
    Set myTest(-1) = mpInc.Deb(1)
    myTest(0) = CByte(42.1)
    myTest(1) = CCur(42.1)
    myTest(2) = CDate(42.1)
    myTest(3) = CDec(42.1)
    myTest(4) = CDbl(42.1)
    myTest(5) = CInt(42.1)
    myTest(6) = CLng(42.1)
    myTest(7) = CLngLng(42.1)
    myTest(8) = CLngPtr(42.1)
    myTest(9) = CSng(42.1)
    myTest(10) = "HelloWorld"
    Set myTest(11) = Seq.Deb
    Dim myResult(0 To 12)  As Variant

    'Act:
    Dim myIndex As Long
    For myIndex = 0 To 12
    myResult(myIndex) = TypeInfo.IsPrimitive(myTest(myIndex - 1))
    Next
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test06b_IsNotPrimitive()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected()  As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, False, False, False, False, True)

    Dim myTest(-1 To 11) As Variant
    
    Set myTest(-1) = mpInc.Deb(1)
    myTest(0) = CByte(42.1)
    myTest(1) = CCur(42.1)
    myTest(2) = CDate(42.1)
    myTest(3) = CDec(42.1)
    myTest(4) = CDbl(42.1)
    myTest(5) = CInt(42.1)
    myTest(6) = CLng(42.1)
    myTest(7) = CLngLng(42.1)
    myTest(8) = CLngPtr(42.1)
    myTest(9) = CSng(42.1)
    myTest(10) = "HelloWorld"
    Set myTest(11) = Seq.Deb
    Dim myResult(0 To 12)  As Variant

    'Act:
    Dim myIndex As Long
    For myIndex = 0 To 12
    myResult(myIndex) = TypeInfo.IsNotPrimitive(myTest(myIndex - 1))
    Next
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test07a_IsItemByIndex0()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, False, False, True, False, False)

    Dim myMapper As mpInc
    Set myMapper = mpInc.Deb(1)

    Dim myResult(0 To 5)  As Boolean

    'Act:
    
    myResult(0) = TypeInfo.IsItemByIndex0(42)
    myResult(1) = TypeInfo.IsItemByIndex0(wCollection.Deb)
    myResult(2) = TypeInfo.IsItemByIndex0(Seq.Deb)
    myResult(3) = TypeInfo.IsItemByIndex0(New ArrayList)
    myResult(4) = TypeInfo.IsItemByIndex0(Queue.Deb)
    myResult(5) = TypeInfo.IsItemByIndex0(Hkvp.Deb)
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test07b_IsNotItemByIndex0()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, True, False, True, True)

    Dim myMapper As mpInc
    Set myMapper = mpInc.Deb(1)

    Dim myResult(0 To 5)  As Boolean

    'Act:
    
    myResult(0) = TypeInfo.IsNotItemByIndex0(42)
    myResult(1) = TypeInfo.IsNotItemByIndex0(wCollection.Deb)
    myResult(2) = TypeInfo.IsNotItemByIndex0(Seq.Deb)
    myResult(3) = TypeInfo.IsNotItemByIndex0(New ArrayList)
    myResult(4) = TypeInfo.IsNotItemByIndex0(Queue.Deb)
    myResult(5) = TypeInfo.IsNotItemByIndex0(Hkvp.Deb)
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TypeInfo")
Private Sub Test08a_IsItemByIndex1()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, True, True, False, False, False)

    Dim myMapper As mpInc
    Set myMapper = mpInc.Deb(1)

    Dim myResult(0 To 5)  As Boolean

    'Act:
    
    myResult(0) = TypeInfo.IsItemByIndex1(42)
    myResult(1) = TypeInfo.IsItemByIndex1(wCollection.Deb)
    myResult(2) = TypeInfo.IsItemByIndex1(Seq.Deb)
    myResult(3) = TypeInfo.IsItemByIndex1(New ArrayList)
    myResult(4) = TypeInfo.IsItemByIndex1(Queue.Deb)
    myResult(5) = TypeInfo.IsItemByIndex1(Hkvp.Deb)
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test08b_IsNotItemByIndex1()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, False, False, True, True, True)

    Dim myMapper As mpInc
    Set myMapper = mpInc.Deb(1)

    Dim myResult(0 To 5)  As Boolean

    'Act:
    
    myResult(0) = TypeInfo.IsNotItemByIndex1(42)
    myResult(1) = TypeInfo.IsNotItemByIndex1(wCollection.Deb)
    myResult(2) = TypeInfo.IsNotItemByIndex1(Seq.Deb)
    myResult(3) = TypeInfo.IsNotItemByIndex1(New ArrayList)
    myResult(4) = TypeInfo.IsNotItemByIndex1(Queue.Deb)
    myResult(5) = TypeInfo.IsNotItemByIndex1(Hkvp.Deb)
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
Private Sub Test09a_IsItemByKey()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, False, False, False, False, True)


    Dim myResult(0 To 5)  As Boolean

    'Act:
    
    myResult(0) = TypeInfo.IsItemByKey(42)
    myResult(1) = TypeInfo.IsItemByKey(wCollection.Deb)
    myResult(2) = TypeInfo.IsItemByKey(Seq.Deb)
    myResult(3) = TypeInfo.IsItemByKey(New ArrayList)
    myResult(4) = TypeInfo.IsItemByKey(Queue.Deb)
    myResult(5) = TypeInfo.IsItemByKey(Hkvp.Deb)
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub Test09b_IsNotItemByKey()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, True, True, True, False)


    Dim myResult(0 To 5)  As Boolean

    'Act:
    
    myResult(0) = TypeInfo.IsNotItemByKey(42)
    myResult(1) = TypeInfo.IsNotItemByKey(wCollection.Deb)
    myResult(2) = TypeInfo.IsNotItemByKey(Seq.Deb)
    myResult(3) = TypeInfo.IsNotItemByKey(New ArrayList)
    myResult(4) = TypeInfo.IsNotItemByKey(Queue.Deb)
    myResult(5) = TypeInfo.IsNotItemByKey(Hkvp.Deb)
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TypeInfo")
Private Sub Test10a_IsItemByArray()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, False, False, False, True, False)


    Dim myResult(0 To 5)  As Boolean

    'Act:
    
    myResult(0) = TypeInfo.IsItemByToArray(42)
    myResult(1) = TypeInfo.IsItemByToArray(wCollection.Deb)
    myResult(2) = TypeInfo.IsItemByToArray(Seq.Deb)
    myResult(3) = TypeInfo.IsItemByToArray(New ArrayList)
    myResult(4) = TypeInfo.IsItemByToArray(Queue.Deb)
    myResult(5) = TypeInfo.IsItemByToArray(Hkvp.Deb)
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub Test10b_IsNotItemByArray()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, True, True, False, True)


    Dim myResult(0 To 5)  As Boolean

    'Act:
    
    myResult(0) = TypeInfo.IsNotItemByArray(42)
    myResult(1) = TypeInfo.IsNotItemByArray(wCollection.Deb)
    myResult(2) = TypeInfo.IsNotItemByArray(Seq.Deb)
    myResult(3) = TypeInfo.IsNotItemByArray(New ArrayList)
    myResult(4) = TypeInfo.IsNotItemByArray(Queue.Deb)
    myResult(5) = TypeInfo.IsNotItemByArray(Hkvp.Deb)
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TypeInfo")
' ' Private Sub Test02c_AreNumbers()

' ' #If twinbasic Then
' '     myProcedureName = CurrentProcedureName
' '     myComponentName = CurrentComponentName
' ' #Else
' '     myProcedureName = ErrEx.LiveCallstack.ProcedureName
' '     myComponentName = ErrEx.LiveCallstack.ModuleName
' ' #End If

' '     'Arrange:
' '     Dim myExpected()  As Variant
' '     myExpected = Array(True, True, True, True, True, True, True, True, True, True)

' '     Dim myTest(0 To 9) As Variant
' '     myTest(0) = CByte(42.1)
' '     myTest(1) = CCur(42.1)
' '     myTest(2) = CDate(42.1)
' '     myTest(3) = CDec(42.1)
' '     myTest(4) = CDbl(42.1)
' '     myTest(5) = CInt(42.1)
' '     myTest(6) = CLng(42.1)
' '     myTest(7) = CLngLng(42.1)
' '     myTest(8) = CLngPtr(42.1) ' vonly in 32 bit 
' '     myTest(9) = CSng(42.1)
   
' '     Dim myResult(0 To 9)  As Boolean

' '     'Act:
' '     myResult(0) = TypeInfo.IsNumber(myTest(0), 42)
' '     myResult(1) = TypeInfo.IsNumber(myTest(1), 42)
' '     myResult(2) = TypeInfo.IsNumber(myTest(2), 42)
' '     myResult(3) = TypeInfo.IsNumber(myTest(3), 42)
' '     myResult(4) = TypeInfo.IsNumber(myTest(4), 42)
' '     myResult(5) = TypeInfo.IsNumber(myTest(5), 42)
' '     myResult(6) = TypeInfo.IsNumber(myTest(6), 42)
' '     myResult(7) = TypeInfo.IsNumber(myTest(7), 42)
    
' '     'LongPtr does not exists as a specific type.  
' '     ' Vartype(LongPtr) returns vbLongLong or vbLong depending on the bitness of VBA
' ' '#If Win64 Then
' '     myResult(8) = TypeInfo.IsNumber(myTest(8), 42)
' ' ' # Else
' '   '   myResult(8) = TypeInfo.IsNumber(myTest(8), NumberType.AsLong)
' ' ' #End If

' '     myResult(9) = TypeInfo.IsNumber(myTest(9), 42)
    
' '     'Assert.Strict:
' '     AssertStrictSequenceEquals myExpected, myResult, myProcedureName

' ' TestExit:
' ' Exit Sub
' ' TestFail:
' ' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' ' Resume TestExit
' ' End Sub

' '@TestMethod("TypeInfo")
' Private Sub Test02d_IsNotNumber()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'     'Arrange:
'     Dim myExpected()  As Variant
'     myExpected = Array(True, True, True, True, True, True, True, True, True, True)

'     Dim myTest(0 To 9) As Variant
'     myTest(0) = "42.1"
'     myTest(1) = "42.1"
'     myTest(2) = "42.1"
'     myTest(3) = "42.1"
'     myTest(4) = "42.1"
'     myTest(5) = "42.1"
'     myTest(6) = "42.1"
'     myTest(7) = "42.1"
'     myTest(8) = "42.1"
'     myTest(9) = "42.1"
   
'     Dim myResult(0 To 9)  As Boolean

'     'Act:
'     myResult(0) = TypeInfo.IsNotANumber(myTest(0), 42)
'     myResult(1) = TypeInfo.IsNotANumber(myTest(1), 42)
'     myResult(2) = TypeInfo.IsNotANumber(myTest(2), 42)
'     myResult(3) = TypeInfo.IsNotANumber(myTest(3), 42)
'     myResult(4) = TypeInfo.IsNotANumber(myTest(4), 42)
'     myResult(5) = TypeInfo.IsNotANumber(myTest(5), 42)
'     myResult(6) = TypeInfo.IsNotANumber(myTest(6), 42)
'     myResult(7) = TypeInfo.IsNotANumber(myTest(7), 42)
    
'     'LongPtr does not exists as a specific type.  
'     ' Vartype(LongPtr) returns vbLongLong or vbLong depending on the bitness of VBA
' '#If Win64 Then
'     myResult(8) = TypeInfo.IsNotANumber(myTest(8), NumberType.IsLongLong)
' '# Else
'   '  myResult(8) = TypeInfo.IsNumber(myTest(8), NumberType.AsLong)
' '#End If

    
'     myResult(9) = TypeInfo.IsNotANumber(myTest(9), NumberType.IsSingle)
    
'     'Assert.Strict:
'     AssertStrictSequenceEquals myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub
' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit
' End Sub



' '@TestMethod("TypeInfo")
' Private Sub Test01c_IsString()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False

'     Dim myTest(0 To 10) As Variant
'     myTest(0) = CByte(42.1)
'     myTest(1) = CCur(42.1)
'     myTest(2) = CDate(42.1)
'     myTest(3) = CDec(42.1)
'     myTest(4) = CDbl(42.1)
'     myTest(5) = CInt(42.1)
'     myTest(6) = CLng(42.1)
'     myTest(7) = CLngLng(42.1)
'     myTest(8) = CLngPtr(42.1)
'     myTest(9) = CSng(42.1)
'     myTest(10) = "HelloWorld"


'     Dim myResult  As Boolean

'     'Act:
   
'     myResult = TypeInfo.IsString(myTest, "HelloWorld")
   
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

' '@TestMethod("TypeInfo")
' Private Sub Test01d_IsNotString()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True

'     Dim myTest(0 To 10) As Variant
'     myTest(0) = CByte(42.1)
'     myTest(1) = CCur(42.1)
'     myTest(2) = CDate(42.1)
'     myTest(3) = CDec(42.1)
'     myTest(4) = CDbl(42.1)
'     myTest(5) = CInt(42.1)
'     myTest(6) = CLng(42.1)
'     myTest(7) = CLngLng(42.1)
'     myTest(8) = CLngPtr(42.1)
'     myTest(9) = CSng(42.1)
'     myTest(10) = "HelloWorld"

'     Dim myResult As Boolean

'     'Act:
    
    
'     myResult = TypeInfo.IsNotString(myTest, "HelloWorld")
    
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

' '@TestMethod("TypeInfo")
' Private Sub Test01ce_IsString()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True

'     Dim myTest(0 To 10) As Variant
'     myTest(0) = CStr(CByte(42.1))
'     myTest(1) = CStr(CCur(42.1))
'     myTest(2) = CStr(CDate(42.1))
'     myTest(3) = CStr(CDec(42.1))
'     myTest(4) = CStr(CDbl(42.1))
'     myTest(5) = CStr(CInt(42.1))
'     myTest(6) = CStr(CLng(42.1))
'     myTest(7) = CStr(CLngLng(42.1))
'     myTest(8) = CStr(CLngPtr(42.1))
'     myTest(9) = CStr(CSng(42.1))
'     myTest(10) = "HelloWorld"


'     Dim myResult  As Boolean

'     'Act:
'     ' IsString needs to detect that myTest has a base type of variant and
'     ' therefore look at each item
'     myResult = TypeInfo.IsString(myTest, "Hello World Again")
   
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub


' '@TestMethod("TypeInfo")
' Private Sub Test02a_AreSameType_Strings_true()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean = True

    

'     Dim myResult  As Boolean

'     'Act:
 
'     myResult = TypeInfo.AreSameType("Hello", "There", "World")
  
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

' '@TestMethod("TypeInfo")
' Private Sub Test02b_AreSameType_StringsAndInteger_False()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean = False

    

'     Dim myResult  As Boolean

'     'Act:
 
'     myResult = TypeInfo.AreSameType("Hello", "There", 42)
  
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub


' '@TestMethod("TypeInfo")
' Private Sub Test02c_AreSameType_ArrayOfString_True()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean = True
'     Dim myResult  As Boolean

'     'Act:
 
'     myResult = TypeInfo.AreSameType(Array("Hello", "There", "World"))
  
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

' '@TestMethod("TypeInfo")
' Private Sub Test02d_AreSameType_MixedArray_False()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean = False
'     Dim myResult  As Boolean

'     'Act:
 
'     myResult = TypeInfo.AreSameType(Array("Hello", "There", 42))
  
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

' '@TestMethod("TypeInfo")
' Private Sub Test03a_IsEmpty_True()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'     'Arrange:
'     Dim myExpected  As Boolean = True
'     Dim myTest As Variant = Array(Empty, Empty, Empty)
'     Dim myResult As Boolean
'     'Act:
'     myResult = TypeInfo.IsEmpty(myTest(0), myTest(2), myTest(1))
    
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub
' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit
' End Sub

' '@TestMethod("TypeInfo")
' Private Sub Test03b_IsEmpty_False()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'     'Arrange:
'     Dim myExpected  As Boolean = False
'     Dim myTest As Variant = Array(Empty, Empty, 10)
'     Dim myResult As Boolean
'     'Act:
'     myResult = TypeInfo.IsEmpty(myTest(0), myTest(2), myTest(1))
    
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub
' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit
' End Sub

' '@TestMethod("TypeInfo")
' Private Sub Test03c_IsNotEmpty_True()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'     'Arrange:
'     Dim myExpected  As Boolean = True
'     Dim myTest As Variant = Array(10, 3.142, "Hello")
'     Dim myResult As Boolean
'     'Act:
'     myResult = TypeInfo.IsNotEmpty(myTest(0), myTest(2), myTest(1))
    
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub
' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit
' End Sub

' '@TestMethod("TypeInfo")
' Private Sub Test03d_IsNotEmpty_False()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'     'Arrange:
'     Dim myExpected  As Boolean = False
'     Dim myTest As Variant = Array(Empty, Empty, 10)
'     Dim myResult As Boolean
'     'Act:
'     myResult = TypeInfo.IsEmpty(myTest(0), myTest(2), myTest(1))
    
'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub
' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit
' End Sub


' '@TestMethod("TypeInfo.Are")
' Private Sub Test03a_IsTypeByTypeId_MultipleIntegers_True()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

' 'On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = True

' Dim myResult  As Boolean

' 'Act:
' myResult = TypeInfo.IsType(idInteger, 42, 100, -10)

' 'Assert.Strict:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub
' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit
' End Sub


' '@TestMethod("TypeInfo.Are")
' Private Sub Test03b_IsTypeByTypeId_MultipleIntegers_OneSingle_False()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    
'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False

'     Dim myTest(0 To 2) As Variant
'     myTest(0) = 42
'     myTest(1) = CSng(42)
'     myTest(2) = -32000

'     Dim myResult  As Boolean

'     'Act:
'     myResult = TypeInfo.IsType(idInteger, myTest(0), myTest(1), myTest(2))

'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

' '@TestMethod("TypeInfo.Are")
' Private Sub Test03c_IsTypeByTypeName_MultipleIntegers_True()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    
'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True

'     Dim myTest(0 To 2) As Variant
'     myTest(0) = 42
'     myTest(1) = 420
'     myTest(2) = -32000

'     Dim myResult  As Boolean

'     'Act:
'     myResult = TypeInfo.IsType("Integer", myTest(0), myTest(1), myTest(2))

'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
' End Sub


' '@TestMethod("TypeInfo.Are")
' Private Sub Test03d_IsTypeByTypeName_MultipleIntegers_OneSingle_False()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    
'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False

'     Dim myTest(0 To 2) As Variant
'     myTest(0) = 42
'     myTest(1) = CSng(420)
'     myTest(2) = -32000

'     Dim myResult  As Boolean

'     'Act:
'     myResult = TypeInfo.IsType("Integer", myTest(0), myTest(1), myTest(2))

'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
' End Sub


' '@TestMethod("Types")
' Private Sub Test03e_IsTypeByIntegerType_MultipleIntegers_True()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    
'    ''On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True

'     Dim myTest(0 To 2) As Variant
'     myTest(0) = 42
'     myTest(1) = 420
'     myTest(2) = -32000

'     Dim myResult  As Boolean

'     'Act:
'     myResult = TypeInfo.IsTypeByItem(42, myTest(0), myTest(1), myTest(2))

'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
' End Sub


' Private Sub Test03f_IsTypeByHkvp_MultipleIntegers_String_False()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False

'     Dim myTest(0 To 2) As Variant
'     myTest(0) = 42
'     myTest(1) = CSng(420)
'     myTest(2) = "42"

'     Dim myResult  As Boolean

'     'Act:
'     myResult = TypeInfo.IsNumber(myTest(0), myTest(1), myTest(2))

'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
' End Sub





' '@TestMethod("TypeInfo")
' Private Sub Test05_AreSameType_MultipleIntegersWithIntegerArrayIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = False

' Dim myTestArray(1 To 5) As Integer
' myTestArray(1) = 10
' myTestArray(2) = 20
' myTestArray(3) = 30
' myTestArray(4) = 40
' myTestArray(5) = 50

' Dim myTest As Integer
' myTest = 42

' Dim myTest3 As Integer
' myTest3 = -32000

' Dim myResult  As Boolean

' 'Act:
' myResult = TypeInfo.AreSameType(myTest, myTestArray, myTest3)

' 'Assert.Strict:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

' End Sub


' '@TestMethod("TypeInfo.Are")
' Private Sub Test06_MultipleIntegersByIsWithNonIntegerIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'    'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False

'     Dim myTestArray(1 To 5) As Integer
'     myTestArray(1) = 10
'     myTestArray(2) = 20
'     myTestArray(3) = 30
'     myTestArray(4) = 40
'     myTestArray(5) = 50

'     Dim myTest As Integer
'     myTest = 42

'     Dim myTest3 As Integer
'     myTest3 = -32000

'     Dim myResult  As Boolean

'     'Act:
'   '  myResult = TypeInfo.AreSameType CInt(42), myTest, Box(myTestArray), myTest3)

'     'Assert.Strict:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName

'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub

' '@TestMethod("TypeInfo.Are")
' Private Sub Test07_MultipleIntegersWithVariantArrayofIntegerIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = False

' Dim myTestArray As Variant
' myTestArray = Array(10, 20, 30, 40, 50)


' Dim myTest As Integer
' myTest = 42

' Dim myTest3 As Integer
' myTest3 = -32000

' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.AreOf(CInt(42), myTest, myTestArray, myTest3)


' 'Assert:

' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub

' ' '@TestMethod("TypeInfo.Are")
' ' Private Sub Test08_MultipleIntegersWithVariantArrayWithVariantArraysOfInteger()
' ''On Error GoTo TestFail

' ' 'Arrange:
' ' Dim myExpected  As Boolean
' ' myExpected = True

' ' Dim myTestarray As Variant
' ' myTestarray = _
' '     Array _
' '     ( _
' '         Array(10%, 20%, 30%, 40%, 50%), _
' '         Array(10%, 20%, 30%, 40%, 50%), _
' '         Array(10%, 20%, 30%, 40%, 50%), _
' '         Array _
' '         ( _
' '             Array(10%, 20%, 30%, 40%, 50%), _
' '             Array(10%, 20%, 30%, 40%, 50%), _
' '             Array(10%, 20%, 30%, 40%, 50%), _
' '             Array(10%, 20%, 30%, 40%, 50%) _
' '          ) _
' '     )


' ' Dim myTest As Integer
' ' myTest = 42

' ' Dim myTest3 As Integer
' ' myTest3 = -32000

' ' Dim myResult  As Boolean

' ' 'Act:
' ' myResult = TypeInfo.Are(sInteger, myTest, myTestarray, myTest3)


' ' 'Assert:

' ' AssertStrictAreEqual myExpected, myResult, myProcedureName

' ' TestExit:
' ' Exit Sub

' ' TestFail:
' ' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' ' Resume TestExit

    
' ' End Sub


' '@TestMethod("TypeInfo.AreNot)
' Private Sub Test11_SingleNonIntegerIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = True


' Dim myTest As Double
' myTest = 42#

' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.AreNot(sInteger, myTest)

' 'Assert.Strict:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub
' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit
' End Sub

' '@TestMethod("TypeInfo.AreNot")
' Private Sub Test12_MultipleIntegersIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = False

' Dim myTest As Integer
' myTest = 42

' Dim myTest2 As Integer
' myTest2 = 100

' Dim myTest3 As Integer
' myTest3 = -32000

' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.AreNot(sInteger, myTest, myTest2, myTest3)

' 'Assert.Strict:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub
' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit
' End Sub

' '@TestMethod("TypeInfo.AreNot")
' Private Sub Test13_MultipleNonIntegersIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = True

' Dim myTest As Double
' myTest = 42

' Dim myTest2 As Single
' myTest2 = 100

' Dim myTest3 As String
' myTest3 = "42"

' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.AreNot(sInteger, myTest, myTest2, myTest3)

' 'Assert.Strict:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub
' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit
' End Sub

' ' '@TestMethod("TypeInfo.Are")
' ' Private Sub Test13_SingleNonIntegersFalse()
' ''On Error GoTo TestFail

' ' 'Arrange:
' ' Dim myExpected  As Boolean
' ' myExpected = False


' ' Dim myTest As String
' ' myTest = "42"

' ' Dim myResult  As Boolean

' ' 'Act:
' ' myResult = TypeInfo.Are(sInteger, myTest)

' ' 'Assert.Strict:
' ' AssertStrictAreEqual myExpected, myResult, myProcedureName

' ' TestExit:
' ' Exit Sub
' ' TestFail:
' ' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' ' Resume TestExit
' ' End Sub


' '@TestMethod("TypeInfo.AreNot")
' Private Sub Test14_MultipleIntegersWithSingleNonIntegerIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = False

' Dim myTest As Integer
' myTest = 42

' Dim myTest2 As Long
' myTest2 = 100

' Dim myTest3 As Integer
' myTest3 = -32000

' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.AreOf(CInt(42), myTest, myTest2, myTest3)

' 'Assert.Strict:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub
' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit
' End Sub

' '@TestMethod("TypeInfo.AreNot")
' Private Sub Test15_MultipleNonIntegersWithNonIntegerArrayIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = True

' Dim myTestArray(1 To 5) As Long
' myTestArray(1) = 10
' myTestArray(2) = 20
' myTestArray(3) = 30
' myTestArray(4) = 40
' myTestArray(5) = 50

' Dim myTest As Long
' myTest = 42

' Dim myTest3 As Long
' myTest3 = -32000

' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.AreNot(sInteger, myTest, myTestArray, myTest3)

' 'Assert.Strict:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

' End Sub


' ' '@TestMethod("TypeInfo.AreNot")
' ' Private Sub Test16_MultipleNonIntegersWithBoxedNonIntegerArrayIsTrue()
' ''On Error GoTo TestFail

' ' 'Arrange:
' ' Dim myExpected  As Boolean
' ' myExpected = True

' ' Dim myTestarray(1 To 5) As Long
' ' myTestarray(1) = 10
' ' myTestarray(2) = 20
' ' myTestarray(3) = 30
' ' myTestarray(4) = 40
' ' myTestarray(5) = 50

' ' Dim myTest As Long
' ' myTest = 42

' ' Dim myTest3 As Long
' ' myTest3 = -32000

' ' Dim myResult  As Boolean

' ' 'Act:
' ' myResult = TypeInfo.AreNot(sInteger, myTest, Box(myTestarray), myTest3)

' ' 'Assert.Strict:
' ' AssertStrictAreEqual myExpected, myResult, myProcedureName

' ' TestExit:
' ' Exit Sub
' ' TestFail:
' ' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' ' Resume TestExit
' ' End Sub

' '@TestMethod("TypeInfo.AreNot")
' Private Sub Test17_MultipleNonIntegersWithVariantArrayofNonIntegerIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = True

' Dim myTestArray As Variant
' myTestArray = Array(10&, 20&, 30&, 40&, 50&)


' Dim myTest As Long
' myTest = 42

' Dim myTest3 As Long
' myTest3 = -32000

' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.AreNot(sInteger, myTest, myTestArray, myTest3)


' 'Assert:

' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub

' '@TestMethod("TypeInfo.AreNot")
' Private Sub Test18_MultipleIntegersWithVariantArrayWithVariantArraysOfInteger()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = True

' Dim myTestArray As Variant
' myTestArray = _
'     Array _
'     ( _
'         Array(10&, 20&, 30&, 40&, 50&), _
'         Array(10&, 20&, 30&, 40&, 50&), _
'         Array(10&, 20&, 30&, 40&, 50&), _
'         Array _
'         ( _
'             Array(10&, 20&, 30&, 40&, 50&), _
'             Array(10&, 20&, 30&, 40&, 50&), _
'             Array(10&, 20&, 30&, 40&, 50&), _
'             Array(10&, 20&, 30&, 40&, 50&) _
'          ) _
'     )


' Dim myTest As Long
' myTest = 42

' Dim myTest3 As Long
' myTest3 = -32000

' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.AreNot(sInteger, myTest, myTestArray, myTest3)


' 'Assert:

' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub


' '@TestMethod("TypeInfo.IsArray")
' Private Sub Test21_SingleArrayIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = True

' Dim myTestArray As Variant
' myTestArray = Array(10&, 20&, 30&, 40&, 50&)

' Dim myResult  As Boolean

' 'Act:
' myResult = TypeInfo.IsAnArray(myTestArray)

' 'Assert:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub

' '@TestMethod("TypeInfo.IsArray")
' Private Sub Test22_MultipleArraysIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = True


' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.IsArray(Array(10, 20, 30, 40, 50)) ', Array(1, 2, 3, 4, 5), Array("Hello", "there", "world "))

' 'Assert:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub


' '@TestMethod("TypeInfo.IsArray")
' Private Sub Test23_IntegerIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = False


' Dim myResult  As Boolean

' 'Act:
' myResult = TypeInfo.IsAnArray(42)

' 'Assert:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub


' '@TestMethod("TypeInfo.IsArray")
' Private Sub Test24_IntegerAndMultipleArraysIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = False


' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.IsArray(Array(10, 20, 30, 40, 50), Array(1, 2, 3, 4, 5), 42, Array("Hello", "there", "world "))

' 'Assert:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub

' '@TestMethod("TypeInfo")
' Private Sub Test31_SingleIntgerIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
'     Dim myTestArray As Variant
'     myTestArray = 42
    
'     Dim myResult  As Boolean
    
'     'Act:
'     myResult = TypeInfo.IsNotAnArray(myTestArray)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
'     End Sub
    
'     '@TestMethod("TypeInfo.IsNotArray")
'     Private Sub Test32_MultipleNonArraysIsTrue()
    
' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
    
'     Dim myResult  As Boolean
    
'     'Act:
'     'myResult = TypeInfo.IsNotArray(42, 3.147, "Hello World")
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
'     End Sub
    
    
'     '@TestMethod("TypeInfo.IsNotArray")
'     Private Sub Test33_VariantArrayIsFalse()
    
' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False
    
    
'     Dim myResult  As Boolean
    
'     'Act:
'     myResult = TypeInfo.IsNotAnArray(Array(10, 20, 30, 40, 50))
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
'     End Sub
    
    
'     '@TestMethod("TypeInfo.IsArray")
'     Private Sub Test34_IntegerAndMultipleArraysIsFalse()
    
' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False
    
    
'     Dim myResult  As Boolean
    
'     'Act:
'     'myResult = TypeInfo.IsArray(Array(10, 20, 30, 40, 50), Array(1, 2, 3, 4, 5), 42, Array("Hello", "there", "world "))
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
'     End Sub
    
'     '@TestMethod("TypeInfo.IsAnObject)
' Private Sub Test41_SingleObjectIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = True

' Dim myTest As Collection
' Set myTest = New Collection

' Dim myResult  As Boolean

' 'Act:
' myResult = TypeInfo.IsNotItemObject(myTest)

' 'Assert:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub

' '@TestMethod("TypeInfo.IsAnObject)
' Private Sub Test42_MultipleObjectsIsTrue() ' Fails when an objectt is a Box because Box is broken

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = True


' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.IsObject(New ArrayList, New Queue, New Collection)

' 'Assert:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub

' '@TestMethod("TypeInfo.IsAnObject)
' Private Sub Test43_IntegerIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = False


' Dim myResult  As Boolean

' 'Act:
' myResult = TypeInfo.IsNotItemObject(42)

' 'Assert:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub


' '@TestMethod("TypeInfo.IsAnObject)
' Private Sub Test44_VariantIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = False

' Dim myTest As Variant
' myTest = 42

' Dim myResult  As Boolean

' 'Act:
' myResult = TypeInfo.IsAnArray(myTest)

' 'Assert:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub

' '@TestMethod("TypeInfo.IsAnObject)
' Private Sub Test45_IntegerWithMultipleObjectsIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

' On Error GoTo TestFail

' 'Arrange:
' Dim myExpected  As Boolean
' myExpected = False

' Dim myTest As Variant
' myTest = 42

' Dim myResult  As Boolean

' 'Act:
' 'myResult = TypeInfo.IsObject(New Collection, New Dictionary, 42, New ArrayList)

' 'Assert:
' AssertStrictAreEqual myExpected, myResult, myProcedureName

' TestExit:
' Exit Sub

' TestFail:
' Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' Resume TestExit

    
' End Sub

' '@TestMethod("TypeInfo.IsNotAnObject)
' Private Sub Test51_SingleIntgerIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
'     Dim myTest As Variant
'     myTest = 42
    
'     Dim myResult  As Boolean
    
'     'Act:
'     myResult = TypeInfo.IsNotItemObject(myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
'     End Sub
    
' '@TestMethod("TypeInfo.IsNotAnObject)
' Private Sub Test52_MultipleNonObjectssIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
    
'     Dim myResult  As Boolean
    
'     'Act:
'    ' myResult = TypeInfo.IsNotObject(42, 3.147, "Hello World")
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub
    
    
' '@TestMethod("TypeInfo.IsNotAnObject)
' Private Sub Test53_VariantisTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
    
'     Dim myResult  As Boolean
    
'     'Act:
'     myResult = TypeInfo.IsNotItemObject(Array(10, 20, 30, 40, 50))
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub


' '@TestMethod("TypeInfo.IsNotAnObject)
' Private Sub Test54_IntegerAndMultipleObjectsIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
    
'     Dim myResult  As Boolean
    
'     'Act:
'     'myResult = TypeInfo.IsNotObject(Array(10, 20, 30, 40, 50), Array(1, 2, 3, 4, 5), 42, Array("Hello", "there", "world "))
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' '@TestMethod("TypeInfo.ArrayIs)
' Private Sub Test61_IntegerArrayIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
    
'     Dim myResult  As Boolean
'     Dim myTest(1 To 5) As Integer
'     myTest(1) = 1
'     myTest(2) = 2
'     'Act:
'     'myResult = TypeInfo.ArrayIs(sInteger, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub


' Private Sub Test62_VariantIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False
    
'     Dim myResult  As Boolean
'     Dim myTest As Variant
'     myTest = 42
    
'     'Act:
'     'myResult = TypeInfo.ArrayIs(sInteger, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' '@TestMethod("TypeInfo.ArrayIs)
' Private Sub Test63_VariantWithIntegerArrayIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
    
'     Dim myResult  As Boolean
'     Dim myTest(1 To 5) As Integer
'     myTest(1) = 1
'     myTest(2) = 2
    
'     Dim myTestVar As Variant
'     myTestVar = myTest
    
'     'Act:
'     'myResult = TypeInfo.ArrayIs(sInteger, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' '@TestMethod("TypeInfo.ArrayIs)
' Private Sub Test64_VariantArrayNotUniformIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False
    
    
'     Dim myResult  As Boolean
'     Dim myTest As Variant
'     myTest = Array(42, 3.142, "Hello World")
    
'     Dim myTestVar As Variant
'     myTestVar = myTest
    
'     'Act:
'     'myResult = TypeInfo.ArrayIs(sInteger, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' '@TestMethod("TypeInfo.ArrayIs)
' Private Sub Test65_UniformVariantArrayOfVariantArrayIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False
    
    
'     Dim myResult  As Boolean
'     Dim myTest As Variant
'     myTest = _
'         Array _
'         ( _
'             Array(1, 2, 3, 4), _
'             Array(10, 20, 30, 40), _
'             Array _
'             ( _
'                 Array(1, 2, 3), _
'                 Array(10, 20, 30, 40) _
'             ) _
'         )
    
   
'     'Act:
'     'myResult = TypeInfo.ArrayIs(sInteger, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub
' '@TestMethod("TypeInfo.ArrayIs)
' Private Sub Test66_NonUniformVariantArrayOfVariantArrayIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False
    
    
'     Dim myResult  As Boolean
'     Dim myTest As Variant
'     myTest = _
'         Array _
'         ( _
'             Array(1, 2, 3, 4), _
'             Array(10, 20, 30, 40), _
'             Array _
'             ( _
'                 Array(1, 2, 3), _
'                 Array(10, 20, 3.142, 40) _
'             ) _
'         )
    
   
'     'Act:
'     'myResult = TypeInfo.ArrayIs(sInteger, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' '@TestMethod("TypeInfo.ArrayIsNot)
' Private Sub Test71_IntegerArrayOfLongIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
    
'     Dim myResult  As Boolean
'     Dim myTest(1 To 5) As Long
'     myTest(1) = 1&
'     myTest(2) = 2&
'     'Act:
'     'myResult = TypeInfo.ArrayIsNotType(sInteger, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' '@TestMethod("TypeInfo.ArrayIsNot)
' Private Sub Test72_VariantOfLongIsFalse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = False
    
'     Dim myResult  As Boolean
'     Dim myTest As Variant
'     myTest = 42
    
'     'Act:
'     'myResult = TypeInfo.ArrayIsNotType(sInteger, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' '@TestMethod("TypeInfo.ArrayIsNot)
' Private Sub Test73_LongVariantWithIntegerArrayIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
    
'     Dim myResult  As Boolean
'     Dim myTest(1 To 5) As Integer
'     myTest(1) = 1
'     myTest(2) = 2
    
'     Dim myTestVar As Variant
'     myTestVar = myTest
    
'     'Act:
'     'myResult = TypeInfo.ArrayIsNotType(sLong, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' '@TestMethod("TypeInfo.ArrayIsNot)
' Private Sub Test74_VariantArrayNotUniformIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
    
'     Dim myResult  As Boolean
'     Dim myTest As Variant
'     myTest = Array(42, 3.142, "Hello World")
    
'     Dim myTestVar As Variant
'     myTestVar = myTest
    
'     'Act:
'     'myResult = TypeInfo.ArrayIsNotType(sInteger, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' '@TestMethod("TypeInfo.ArrayIsNot)
' Private Sub Test75_UniformVariantArrayOfVariantArrayIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
    
'     Dim myResult  As Boolean
'     Dim myTest As Variant
'     myTest = _
'         Array _
'         ( _
'             Array(1, 2, 3, 4), _
'             Array(10, 20, 30, 40), _
'             Array _
'             ( _
'                 Array(1, 2, 3), _
'                 Array(10, 20, 30, 40) _
'             ) _
'         )
    
   
'     'Act:
'     'myResult = TypeInfo.ArrayIsNotType(sLong, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' '@TestMethod("TypeInfo.ArrayIsNot)
' Private Sub Test76_NonUniformVariantArrayOfVariantArrayIsTrue()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

    

'    'On Error GoTo TestFail
    
'     'Arrange:
'     Dim myExpected  As Boolean
'     myExpected = True
    
'     Dim myResult  As Boolean
'     Dim myTest As Variant
'     myTest = _
'         Array _
'         ( _
'             Array(1, 2, 3, 4), _
'             Array(10, 20, 30, 40), _
'             Array _
'             ( _
'                 Array(1, 2, 3), _
'                 Array(10, 20, 3.142, 40) _
'             ) _
'         )
   
'     'Act:
'     'myResult = TypeInfo.ArrayIsNotType(sInteger, myTest)
    
'     'Assert:
'     AssertStrictAreEqual myExpected, myResult, myProcedureName
    
' TestExit:
'     Exit Sub
    
' TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
    
        
' End Sub

' ''#End Region
