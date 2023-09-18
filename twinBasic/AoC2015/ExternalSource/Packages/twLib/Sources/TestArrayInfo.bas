Attribute VB_Name = "TestArrayInfo"
Option Explicit
Option Private Module
'@IgnoreModule
'@TestModule
'@Folder("Tests")

#If twinbasic Then
    'Do nothing
#Else
'
'
    
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

Public Sub ArrayInfoTests()

    
#If twinbasic Then

    Debug.Print CurrentProcedureName ; vbTab, vbTab,
    
#Else

    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
    
#End If

    Test01a_StaticArrayOfLongIsAllocatedIsTrue
    Test01b_LongIsAllocatedIsFalse
    Test01c_VariantIsAllocatedIsFalse
    Test01d_ReDimmedVariantIsAllocatedIsTrue
    Test01e_DynamicArrayOfLongWithNoDimensionsIsAllocatedIsFalse
    Test01f_DynamicArrayOfLongWithNoDImensionsIsNotAllocatedIsTrue
    
    Test02a_DynamicArrayOfLongWithNoDImensionsHasNoItemsIsTrue
    Test02b_DynamicArrayOfLongWithNoDimensionsHasOneItemIsFalse
    Test02c_DynamicArrayOfLongWithNoDimensionsHasItemsIsFalse
    Test02d_DynamicArrayOfLongWithNoDimensionsHasAnyItemsIsFalse
    Test02e_DynamicArrayOfLongWithFiveItemsHasNoItemsIsFalse
    Test02f_DynamicArrayOfLongWithFiveItemsHasOneItemIsFalse
    Test02g_DynamicArrayOfLongWithFiveItemsHasItemsIsTrue
    Test02h_DynamicArrayOfLongWithFiveItemsHasAnyItemsIsTrue
    Test02i_DynamicArrayOfArrayOfFiveItemsHasNoItemsIsFalse
    Test02j_DynamicArrayOfArrayWithFiveItemsHasOneItemIsTrue
    Test02k_DynamicArrayOfArrayWithFiveItemsHasItemsIsFalse
    Test02l_DynamicArrayOfArrayWithFiveItemsHasAnyItemsIsTrue
    
    Test03a_ArrayOfLongWithNoDimensionsRanksIsZero
    Test03b_ArrayOfLongWithOneDimensionRanksIsOne
    Test03c_ArrayOfLongWithTwoDimensionsRanksIsTwo
    Test03d_ArrayOfLongWithThreeDimensionsRanksIsThree
    Test03e_ArrayOfLongWithNoDimensionsHasRankTwoIsFalse
    Test03f_ArrayOfLongWithOneDimensionHasRankTwoIsFalse
    Test03g_ArrayOfLongWithTwoDimensionsHasRankTwoIsTrue
    Test03h_ArryOfLongWithTwoDimensionsLacksRankTwoIsFalse
    Test03j_ArrayOfLongWithMultiDimensionsHasRankTwoIsTrue
    Test03k_ArrayOfLongWithMultiDimensionsLacksRankTwoIsFalse

    Test04a_ArrayOfLongWithNoDimensionsCountIsZero
    Test04b_ArrayOfLongWithFiveItemsCountIsFive
    Test04c_ArrayOfLongWithTwoByFiveItemsCountIsTwentyFive
    Test04d_ArrayOfLongWithFourByFiveItemsCountIsSixTwoFive
    Test04e_ArrayOfLongWithDimsSizesOfThreeFourFiveSixItemsCountoOfRankOneIsThree
    Test04f_ArrayOfLongWithDimsSizesOfThreeFourFiveSixItemsCountoOfRankTwoIsFour
    Test04g_ArrayOfLongWithDimsSizesOfThreeFourFiveSixItemsCountoOfRankThreeIsFive
    Test04h_ArrayOfLongWithDimsSizesOfThreeFourFiveSixItemsCountoOfRankFourIsSix
    Test04i_ArraysWithSameDimensionsIsSameCountIsTrue
    Test04j_ArraysWithSameDimensionsRankTwoIsSameCountIsTrue
    Test04k_ArraysOfSameSizeWithDifferentRanksIsSameCountIsTrue
    Test04l_DifferentRanksAreSameSizeIsSameCountIsTrue
    Test04m_RanksAreNottSameSizeIsSameCountIsFalse
    
    Test05a_ArrayOfLongBaseTypeIsLong
    Test05b_ArrayOfVariantBaseTypeIsVariant
    Test05c_ArrayOfLongAssignedToVarianttBaseTypeIsLong
    
    Test06a_ArrayofLongWithZeroDimensionsIsListArrayIsFalse
    Test06b_ArrayOfLongWithOneDImesionIsListArrayIsTrue
    Test06c_ArrayOfLongWithTwoDImensionsIsListArrayIsFalse
    Test06d_ArrayOfLongWithMultiDimensionsIsListArrayIsFalse
    Test06e_ArrayofLongWithZeroDimensionsIsTableArrayIsFalse
    Test06f_ArrayOfLongWithOneDImesionIsTableArrayIsFalse
    Test06g_ArrayOfLongWithTwoDImensionsIsTableArrayIsTrue
    Test06h_ArrayOfLongWithMultiDimensionsIsTableArrayIsFalse
    Test06i_ArrayofLongWithZeroDimensionsIsMDArrayIsFalse
    Test06j_ArrayOfLongWithOneDImesionIsMDArrayIsFalse
    Test06k_ArrayOfLongWithTwoDImensionsIsMDArrayIsFalse
    Test06l_ArrayOfLongWithMultiDimensionsIsMDArrayIsTrue
    
    Debug.Print "Testing completed"

End Sub
    

Public Function MakeRowColArray(ByVal ipRows As Long, ByVal ipCols As Long) As Variant

    '@Ignore VariableNotAssigned
    Dim myArray As Variant
    ReDim myArray(1 To ipRows, 1 To ipCols)
    Dim myValue As Long
    myValue = 1
    
    Dim myRow As Long
    For myRow = 1 To ipRows
    
        Dim myCol As Long
        For myCol = 1 To ipCols
        
            myArray(myRow, myCol) = myValue
            myValue = myValue + 1
            
        Next
        
    Next
        
    MakeRowColArray = myArray
    
End Function


Public Function MakeColRowArray(ByVal ipRows As Long, ByVal ipCols As Long) As Variant

    '@Ignore VariableNotAssigned
    Dim myArray As Variant
    ReDim myArray(1 To ipRows, 1 To ipCols)
    Dim myValue As Long
    myValue = 1
    
    Dim myCol As Long
    For myCol = 1 To ipCols
    
        Dim myRow As Long
        For myRow = 1 To ipRows
        
            myArray(myRow, myCol) = myValue
            myValue = myValue + 1
            
        Next
        
    Next
        
    MakeColRowArray = myArray
    
End Function

Public Function GetParamArray(ParamArray ipArgs() As Variant) As Variant
    GetParamArray = ipArgs
End Function



'@TestMethod("Arrays")
Public Sub Test01a_StaticArrayOfLongIsAllocatedIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    '''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 5) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsAllocated(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test01b_LongIsAllocatedIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
    '''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray As Long ' Not a long array, just a long
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsAllocated(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub




'@TestMethod("Arrays")
Public Sub Test01c_VariantIsAllocatedIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray As Variant ' not an array, just variant
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsAllocated(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test01d_ReDimmedVariantIsAllocatedIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If


    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray As Variant
    ReDim myArray(1 To 5)
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsAllocated(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test01e_DynamicArrayOfLongWithNoDimensionsIsAllocatedIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If


    '''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.IsAllocated(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test01f_DynamicArrayOfLongWithNoDImensionsIsNotAllocatedIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If


    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray() As Long
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.IsNotAllocated(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02a_DynamicArrayOfLongWithNoDImensionsHasNoItemsIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If


    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray() As Long
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.IsNotQueryable(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02b_DynamicArrayOfLongWithNoDimensionsHasOneItemIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.HasOneItem(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02c_DynamicArrayOfLongWithNoDimensionsHasItemsIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.HasItems(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02d_DynamicArrayOfLongWithNoDimensionsHasAnyItemsIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.HasAnyItems(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02e_DynamicArrayOfLongWithFiveItemsHasNoItemsIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 5) As Long
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.IsNotQueryable(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02f_DynamicArrayOfLongWithFiveItemsHasOneItemIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.HasOneItem(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02g_DynamicArrayOfLongWithFiveItemsHasItemsIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 5) As Long
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.HasItems(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02h_DynamicArrayOfLongWithFiveItemsHasAnyItemsIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 5) As Long
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.HasAnyItems(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Arrays")
Public Sub Test02i_DynamicArrayOfArrayOfFiveItemsHasNoItemsIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray As Variant
    myArray = Array(Array(1, 2, 3, 4, 5))
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.IsNotQueryable(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02j_DynamicArrayOfArrayWithFiveItemsHasOneItemIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray As Variant
    myArray = Array(Array(1, 2, 3, 4, 5))
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.HasOneItem(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02k_DynamicArrayOfArrayWithFiveItemsHasItemsIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray As Variant
    myArray = Array(Array(1, 2, 3, 4, 5))
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.HasItems(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test02l_DynamicArrayOfArrayWithFiveItemsHasAnyItemsIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray As Variant
    myArray = Array(Array(1, 2, 3, 4, 5))
    Dim myResult As Boolean
  
    'Act:
    myResult = ArrayInfo.HasAnyItems(myArray)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Arrays")
Public Sub Test03a_ArrayOfLongWithNoDimensionsRanksIsZero()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 0
    
    Dim myArray() As Long
    Dim myResult As Long
   
    'Act:
    myResult = ArrayInfo.Ranks(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test03b_ArrayOfLongWithOneDimensionRanksIsOne()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 1
    
    Dim myArray(1 To 5) As Long
    Dim myResult As Long
   
    'Act:
    myResult = ArrayInfo.Ranks(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test03c_ArrayOfLongWithTwoDimensionsRanksIsTwo()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 2
    
    Dim myArray(1 To 10, 1 To 10) As Long
    Dim myResult As Long
   
    'Act:
    myResult = ArrayInfo.Ranks(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Arrays")
Public Sub Test03d_ArrayOfLongWithThreeDimensionsRanksIsThree()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 3
    
    Dim myArray(1 To 10, 1 To 10, 1 To 10) As Long
    Dim myResult As Long
   
    'Act:
    myResult = ArrayInfo.Ranks(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("HasRank")
Public Sub Test03e_ArrayOfLongWithNoDimensionsHasRankTwoIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.HasRank(myArray, 2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("HasRank")
Public Sub Test03f_ArrayOfLongWithOneDimensionHasRankTwoIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 10) As Long
    Dim myResult As Boolean
   
    'Act:
    myResult = ArrayInfo.HasRank(myArray, 2)
    

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("HasRank")
Public Sub Test03g_ArrayOfLongWithTwoDimensionsHasRankTwoIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 10, 1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.HasRank(myArray, 2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("HasRank")
Public Sub Test03h_ArryOfLongWithTwoDimensionsLacksRankTwoIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.HasRank(myArray, 2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test03j_ArrayOfLongWithMultiDimensionsHasRankTwoIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 10, 1 To 10, 1 To 10, 1 To 10) As Long
    Dim myResult As Boolean
   
    'Act:
    myResult = ArrayInfo.HasRank(myArray, 2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Arrays")
Public Sub Test03k_ArrayOfLongWithMultiDimensionsLacksRankTwoIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 10, 1 To 10, 1 To 10, 1 To 10) As Long
    Dim myResult As Boolean
   
    'Act:
    myResult = ArrayInfo.LacksRank(myArray, 2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub




'@TestMethod("Arrays")
Public Sub Test04a_ArrayOfLongWithNoDimensionsCountIsZero()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = -1
    
    Dim myArray() As Long
    Dim myResult As Long
    
    'Act:
    myResult = ArrayInfo.Count(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test04b_ArrayOfLongWithFiveItemsCountIsFive()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 5
    
    Dim myArray(1 To 5) As Long
    Dim myResult As Long
    
    'Act:
    myResult = ArrayInfo.Count(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test04c_ArrayOfLongWithTwoByFiveItemsCountIsTwentyFive()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 25
    
    Dim myArray(1 To 5, 1 To 5) As Long
    Dim myResult As Long
    
    'Act:
    myResult = ArrayInfo.Count(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test04d_ArrayOfLongWithFourByFiveItemsCountIsSixTwoFive()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 625
    
    Dim myArray(1 To 5, 1 To 5, 1 To 5, 1 To 5) As Long
    Dim myResult As Long
    
    'Act:
    myResult = ArrayInfo.Count(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Arrays")
Public Sub Test04e_ArrayOfLongWithDimsSizesOfThreeFourFiveSixItemsCountoOfRankOneIsThree()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 3
    
    Dim myArray(1 To 3, 1 To 4, 1 To 5, 1 To 6) As Long
    Dim myResult As Long
    
    'Act:
    myResult = ArrayInfo.Count(myArray, 1)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub
'@TestMethod("Arrays")
Public Sub Test04f_ArrayOfLongWithDimsSizesOfThreeFourFiveSixItemsCountoOfRankTwoIsFour()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 4
    
    Dim myArray(1 To 3, 1 To 4, 1 To 5, 1 To 6) As Long
    Dim myResult As Long
    
    'Act:
    myResult = ArrayInfo.Count(myArray, 2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub
'@TestMethod("Arrays")
Public Sub Test04g_ArrayOfLongWithDimsSizesOfThreeFourFiveSixItemsCountoOfRankThreeIsFive()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 5
    
    Dim myArray(1 To 3, 1 To 4, 1 To 5, 1 To 6) As Long
    Dim myResult As Long
    
    'Act:
    myResult = ArrayInfo.Count(myArray, 3)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test04h_ArrayOfLongWithDimsSizesOfThreeFourFiveSixItemsCountoOfRankFourIsSix()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 6
    
    Dim myArray(1 To 3, 1 To 4, 1 To 5, 1 To 6) As Long
    Dim myResult As Long
    
    'Act:
    myResult = ArrayInfo.Count(myArray, 4)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Arrays")
Public Sub Test04i_ArraysWithSameDimensionsIsSameCountIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 5, 1 To 2) As Long
    Dim myResult As Boolean
    Dim myTest(1 To 5, 1 To 2) As Long
    'Act:
    myResult = ArrayInfo.IsSameCount(myArray, myTest)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test04j_ArraysWithSameDimensionsRankTwoIsSameCountIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 5, 1 To 2) As Long
    Dim myResult As Boolean
    Dim myTest(1 To 5, 1 To 2) As Long
    'Act:
    myResult = ArrayInfo.IsSameCount(myArray, myTest, 2, 2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Arrays")
Public Sub Test04k_ArraysOfSameSizeWithDifferentRanksIsSameCountIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 5, 1 To 2) As Long
    Dim myResult As Boolean
    Dim myTest(1 To 2, 1 To 5) As Long
    'Act:
    myResult = ArrayInfo.IsSameCount(myArray, myTest)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test04l_DifferentRanksAreSameSizeIsSameCountIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 5, 1 To 2) As Long
    Dim myResult As Boolean
    Dim myTest(1 To 2, 1 To 5) As Long
    'Act:
    myResult = ArrayInfo.IsSameCount(myArray, myTest, 1, 2)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test04m_RanksAreNottSameSizeIsSameCountIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 5, 1 To 2) As Long
    Dim myResult As Boolean
    Dim myTest(1 To 2, 1 To 5) As Long
    'Act:
    myResult = ArrayInfo.IsSameCount(myArray, myTest, 1, 1)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test05a_ArrayOfLongBaseTypeIsLong()

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
    myExpected = "long"
    
    Dim myArray(1 To 5, 1 To 2) As Long
    Dim myResult As String
  
    'Act:
    myResult = TypeInfo.BaseType(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Arrays")
Public Sub Test05b_ArrayOfVariantBaseTypeIsVariant()

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
    myExpected = "variant"
    
    Dim myArray As Variant
    myArray = Array(1&, 2&, 3&, 4&, 57)
    Dim myResult As String
  
    'Act:
    myResult = TypeInfo.BaseType(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test05c_ArrayOfLongAssignedToVarianttBaseTypeIsLong()

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
    myExpected = "long"
    
    Dim myArray As Variant
    Dim myTest(1 To 5) As Long
    myArray = myTest
    Dim myResult As String
  
    'Act:
    myResult = TypeInfo.BaseType(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06a_ArrayofLongWithZeroDimensionsIsListArrayIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsListArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06b_ArrayOfLongWithOneDImesionIsListArrayIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 5) As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsListArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06c_ArrayOfLongWithTwoDImensionsIsListArrayIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 5, 1 To 5) As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsListArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06d_ArrayOfLongWithMultiDimensionsIsListArrayIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 5, 1 To 5, 1 To 5, 1 To 5) As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsListArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06e_ArrayofLongWithZeroDimensionsIsTableArrayIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsListArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06f_ArrayOfLongWithOneDImesionIsTableArrayIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 5) As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsTableArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06g_ArrayOfLongWithTwoDImensionsIsTableArrayIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 5, 1 To 5) As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsTableArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06h_ArrayOfLongWithMultiDimensionsIsTableArrayIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 5, 1 To 5, 1 To 5, 1 To 5) As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsTableArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub



'@TestMethod("Arrays")
Public Sub Test06i_ArrayofLongWithZeroDimensionsIsMDArrayIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsMDArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06j_ArrayOfLongWithOneDImesionIsMDArrayIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 5) As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsMDArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06k_ArrayOfLongWithTwoDImensionsIsMDArrayIsFalse()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 5, 1 To 5) As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsMDArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Arrays")
Public Sub Test06l_ArrayOfLongWithMultiDimensionsIsMDArrayIsTrue()

    #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If

    ''On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 5, 1 To 5, 1 To 5, 1 To 5) As Long
    
    Dim myResult As Boolean
    
    'Act:
    myResult = ArrayInfo.IsMDArray(myArray)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


' The following tests need to me moved to testing of the 
' a different class as they are about array manipulation
' rather than array meta data


' ' '@TestMethod("TryRotate")
' '     Public Sub Test26_TransposeArray()
' '         ''On Error GoTo TestFail
    
' '         'Arrange:
    
' '         Dim myExpectedStatus As Boolean
' '         myExpectedStatus = True
    
' '         Dim myexpectedarray As Variant
' '         myexpectedarray = MakeColRowArray(5, 4)
    
' '         Dim mySource As Variant
' '         mySource = MakeRowColArray(4, 5)
    

' '         Dim myResult As Variant
    
    
' '         'Act:
' '         myResult = ArrayInfo.Transpose(mySource)
    
' '         'Assert:
    
' '         AssertStrictSequenceEquals myexpectedarray, myResult, "Value"
    
    
' ' TestExit:
' '         Exit Sub
    
' ' TestFail:
' '         Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' '         Resume TestExit
    
' '     End Sub

' ' '@TestMethod("TryToListOfList")
' '     Public Sub Test27_ArrayToListOfListsByRow()
' '         ''On Error GoTo TestFail
    
' '         'Arrange:
' '         Dim myExpectedStatus As Boolean
' '         myExpectedStatus = True
    
' '         Dim myExpectedList As List
' '         Set myExpectedList = List.Deb
    
' '         ' This is an example of the idiosyncrasy introduced by ParseVariantUsingSingleItemSpecialCase
' '         ' Which is that if we wish to add a single iterable to a List as a single item
' '         ' the single iterable must be encapsulated in an array.
    
' '         With myExpectedList
    
' '             .Add List.Deb.Add(1&, 2&, 3&, 4&)
' '             .Add List.Deb.Add(5&, 6&, 7&, 8&)
' '             .Add List.Deb.Add(9&, 10&, 11&, 12&)
' '             .Add List.Deb.Add(13&, 14&, 15&, 16&)
' '             .Add List.Deb.Add(17&, 18&, 19&, 20&)
        
' '         End With
    
' '         Dim mySource As Variant
' '         mySource = MakeRowColArray(5, 4)
    
' '         Dim myResult As List
    
' '         'Act:
' '         Set myResult = ArrayInfo.ToEnumerableOfRanksAsEnumerable(mySource, RankIsRowFirstItemActionIsNoAction)
    
' '         'Assert:
    
' '         Dim myIndex As Long
' '         For myIndex = 1 To 5
    
' '             AssertStrictSequenceEquals myExpectedList.Item(myIndex).ToArray, myResult.Item(myIndex).ToArray, myProcedureName
' '         Next
    
' ' TestExit:
' '         Exit Sub
    
' ' TestFail:
' '         Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' '         Resume TestExit
' '     End Sub

' ' '@TestMethod("TryToListOfList")
' '     Public Sub Test28_ArrayToListOfListsByCol()
' '         ''On Error GoTo TestFail
    
' '         'Arrange:
' '         Dim myExpectedStatus As Boolean
' '         myExpectedStatus = True
    
' '         Dim myExpectedList As List
' '         Set myExpectedList = List.Deb
    
' '         With myExpectedList
        
' '             .Add List.Deb.Add(1&, 5&, 9&, 13&, 17&)
' '             .Add List.Deb.Add(2&, 6&, 10&, 14&, 18&)
' '             .Add List.Deb.Add(3&, 7&, 11&, 15&, 19&)
' '             .Add List.Deb.Add(4&, 8&, 12&, 16&, 20&)
        
' '         End With
    
' '         Dim mySource As Variant
' '         mySource = MakeRowColArray(5, 4)
    
' '         Dim myResult As List
    
' '         'Act:
' '         Set myResult = ArrayInfo.ToEnumerableOfRanksAsEnumerable(mySource, TableSlicer.RankIsColumnFirstItemActionIsNoAction)
    
' '         'Assert:

' '         Dim myIndex As Long
' '         For myIndex = 1 To 4
' '             AssertStrictSequenceEquals myExpectedList.Item(myIndex).ToArray, myResult.Item(myIndex).ToArray, myProcedureName
' '         Next
    
' ' TestExit:
' '         Exit Sub
    
' ' TestFail:
' '         Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' '         Resume TestExit
' '     End Sub

' ' '@TestMethod("TryToListOfList")
' '     Public Sub Test29_ArrayToListOfListsByRowFirstItemActionIsSplitFirstRow()
' '         ''On Error GoTo TestFail
    
' '         'Arrange:
' '         Dim myExpectedStatus As Boolean
' '         myExpectedStatus = True
    
' '         Dim myExpectedList As List
' '         Set myExpectedList = List.Deb
    
' '         With myExpectedList
    
' '             Dim myFirstValues As List
' '             Set myFirstValues = List.Deb.AddRange(Array(1&, 5&, 9&, 13&, 17&))
        
' '             .Add myFirstValues
        
' '             Dim myRankValues As List
' '             Set myRankValues = List.Deb
    
' '             With myRankValues
        
' '                 .Add List.Deb.Add(2&, 3&, 4&)
' '                 .Add List.Deb.Add(6&, 7&, 8&)
' '                 .Add List.Deb.Add(10&, 11&, 12&)
' '                 .Add List.Deb.Add(14&, 15&, 16&)
' '                 .Add List.Deb.Add(18&, 19&, 20&)
        
' '             End With
    
' '             .Add myRankValues
        
' '         End With
    
' '         Dim mySource As Variant
' '         mySource = MakeRowColArray(5, 4)
    
' '         Dim myResult As List
    
' '         'Act:
' '         Set myResult = ArrayInfo.ToEnumerableOfRanksAsEnumerable(mySource, RankIsRowFirstItemActionIsSplit)
    
' '         'Assert:
' '         AssertStrictSequenceEquals myFirstValues.ToArray, myResult.First.ToArray, myProcedureName

' '         Dim myIndex As Long
' '         For myIndex = 1 To 5
' '             ' Dim myE As
' '             ' myE = myRankValues.Item(myIndex).toarray
' '             AssertStrictSequenceEquals myExpectedList.Item(2).Item(myIndex).ToArray, myResult.Item(2).Item(myIndex).ToArray, myProcedureName
' '         Next
    
' ' TestExit:
' '         Exit Sub
    
' ' TestFail:
' '         Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' '         Resume TestExit
' '     End Sub

' ' '@TestMethod("TryToListOfList")
' '     Public Sub Test30_ArrayToListOfListsByColFirstItemActionIsSplitFirstItem()
' '         ''On Error GoTo TestFail
    
' '         'Arrange:
' '         Dim myExpectedStatus As Boolean
' '         myExpectedStatus = True
    
' '         Dim myExpectedList As List
' '         Set myExpectedList = List.Deb
    
' '         With myExpectedList
    
' '             Dim myFirstValues As List
' '             Set myFirstValues = List.Deb.Add(1&, 2&, 3&, 4&)
        
' '             .Add myFirstValues
        
' '             Dim myRankValues As List
' '             Set myRankValues = List.Deb
    
' '             With myRankValues
        
' '                 .Add List.Deb.Add(5&, 9&, 13&, 17&)
' '                 .Add List.Deb.Add(6&, 10&, 14&, 18&)
' '                 .Add List.Deb.Add(7&, 11&, 15&, 19&)
' '                 .Add List.Deb.Add(8&, 12&, 16&, 20&)
        
' '             End With
    
' '             .Add myRankValues

    
' '         End With
    
' '         Dim mySource As Variant
' '         mySource = MakeRowColArray(5, 4)
    
' '         Dim myResult As List
    
' '         'Act:
    
' '         Set myResult = ArrayInfo.ToEnumerableOfRanksAsEnumerable(mySource, RankIsColumnFirstItemActionIsSplit)
    
' '         'Assert:
' '         Dim myexpectedarray As Variant
' '         myexpectedarray = myFirstValues.ToArray
    
' '         Dim myResultarray As Variant
' '         myResultarray = myResult.First.ToArray
    
' '         AssertStrictSequenceEquals myFirstValues.ToArray, myResult.First.ToArray, myProcedureName
' '         Dim myIndex As Long
' '         For myIndex = 1 To 4
' '             AssertStrictSequenceEquals myRankValues.Item(myIndex).ToArray, myResult.Item(2).Item(myIndex).ToArray, myProcedureName
' '         Next
    
' ' TestExit:
' '         Exit Sub
    
' ' TestFail:
' '         Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' '         Resume TestExit
' '     End Sub


' ' '@TestMethod("TryToListOfList")
' '     Public Sub Test31_ArrayToListOfListsByRowFirstItemActionIsCopyFirstItem()
' '         ''On Error GoTo TestFail
    
' '         'Arrange:
' '         Dim myExpectedStatus As Boolean
' '         myExpectedStatus = True
    
' '         ' Dim myExpectedList As List
' '         ' Set myExpectedList = List.Deb
    
' '         'With myExpectedList
    
' '             Dim myFirstValues As Variant
' '             myFirstValues = Array(1&, 5&, 9&, 13&, 17&)
        
' '         ' .Add myFirstValues
        
' '             Dim myRankValues As List
' '             Set myRankValues = List.Deb
    
' '             With myRankValues
        
' '                 .Add List.Deb.Add(1&, 2&, 3&, 4&)
' '                 .Add List.Deb.Add(5&, 6&, 7&, 8&)
' '                 .Add List.Deb.Add(9&, 10&, 11&, 12&)
' '                 .Add List.Deb.Add(13&, 14&, 15&, 16&)
' '                 .Add List.Deb.Add(17&, 18&, 19&, 20&)
        
' '             End With
    
' '             '.Add myRankValues
        
' '     ' End With
    
' '         Dim mySource As Variant
' '         mySource = MakeRowColArray(5, 4)
    
' '         Dim myResult As List
    
' '         'Act:
' '         Set myResult = ArrayInfo.ToEnumerableOfRanksAsEnumerable(mySource, RankIsRowFirstItemActionIsCopy)
    
' '         'Assert:
' '         AssertStrictSequenceEquals myFirstValues, myResult.First.ToArray, myProcedureName
' '         Dim myIndex As Long
' '         For myIndex = 1 To 5
' '             AssertStrictSequenceEquals myRankValues.Item(myIndex).ToArray, myResult.Item(2).Item(myIndex).ToArray, myProcedureName
' '         Next
    
' ' TestExit:
' '         Exit Sub
    
' ' TestFail:
' '         Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' '         Resume TestExit
' '     End Sub


' ' '@TestMethod("TryToListOfList")
' '     Public Sub Test32_ArrayToListOfListsByColFirstActionItemIsSPlitFirstItem()
' '         ''On Error GoTo TestFail
    
' '         'Arrange:
' '         Dim myExpectedStatus As Boolean
' '         myExpectedStatus = True
    
' '         ' Dim myExpectedList As List
' '         ' Set myExpectedList = List.Deb
    
' '     '  With myExpectedList
    
' '             Dim myFirstValues As Variant
' '             myFirstValues = Array(1&, 2&, 3&, 4&)
        
' '         '    .Add myFirstValues
        
' '             Dim myRankValues As List
' '             Set myRankValues = List.Deb
    
' '             With myRankValues
        
' '                 .Add List.Deb.Add(5&, 9&, 13&, 17&)
' '                 .Add List.Deb.Add(6&, 10&, 14&, 18&)
' '                 .Add List.Deb.Add(7&, 11&, 15&, 19&)
' '                 .Add List.Deb.Add(8&, 12&, 16&, 20&)
        
        
' '             End With
    
' '         '   .Add myRankValues

    
' '     '  End With
    
' '         Dim mySource As Variant
' '         mySource = MakeRowColArray(5, 4)
    
' '         Dim myResult As List
    
' '         'Act:
' '         Set myResult = ArrayInfo.ToEnumerableOfRanksAsEnumerable(mySource, RankIsColumnFirstItemActionIsSplit)
    
' '         'Assert:
' '         AssertStrictSequenceEquals myFirstValues, myResult.First.ToArray, myProcedureName
' '         Dim myIndex As Long
' '         For myIndex = 1 To 4
' '             AssertStrictSequenceEquals myRankValues.Item(myIndex).ToArray, myResult.Item(2).Item(myIndex).ToArray, myProcedureName
' '         Next
    
' ' TestExit:
' '         Exit Sub
    
' ' TestFail:
' '         Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
' '         Resume TestExit
' '     End Sub
