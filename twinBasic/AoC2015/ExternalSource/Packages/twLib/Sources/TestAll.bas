Attribute VB_Name = "AllTesting"
'@IgnoreModule
'@TestModule
'@Folder("Tests")

Public myProcedureName As String
Public myComponentName As String

#If twinbasic Then
    'currently do nothing
#Else
    Public Assert As Object
    Public Fakes As Object
    ErrEx.Enable vbNullString
#End If

Public Sub AllTests()
    
    Debug.Print "Testing started"
    
    TestArrayInfo.ArrayInfoTests                    ' Pass
    TestSeq.SeqTests                                ' Pass
    TestHkvp.HkvpTests                              ' pass
    TestVariantParser.VariantParserTests            ' Pass
    TestStrs.StrsTests                              ' Pass
    TestIterNum.IterNumTests                        ' Pass
    TestIterItems.IterItemsTest                     ' Pass
    TestRank.RankTests                              ' Pass
    TestExtent.ExtentTests                          ' Pass
    TestStringifier.StringifierTests                ' Pass
    TestFmt.FmtTests                                ' Pass

    ' TestTypeInfo.TypeInfoTests                        ' Pass
    'TestResult.ResultTests                      ' Fail
    
  
    'TestUnsafe.UnsafeTests                      ' Pass
   
   
    'TestRanges.RangeTests

    
    Debug.Print "Testing completed"
    
End Sub


