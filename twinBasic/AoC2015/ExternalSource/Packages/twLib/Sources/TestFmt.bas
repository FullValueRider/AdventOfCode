Attribute VB_Name = "TestFmt"
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


Public Sub FmtTests()

#If twinbasic Then

    Debug.Print CurrentProcedureName ; vbTab, vbTab, vbTab,
    
#Else

    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
    
#End If
    
    T01_ConvertFormatFieldXX0ToVbNullString_01
    T02_ConvertFormatFieldXXToXX1_01
    T03_GetFieldCount_01_nl3_equals_3
    T04_GetFormattingReplaceString_01_nl3
    T05_GetFormattingReplaceString_02_tb4
    T06_GetFormattingReplaceString_03_nt2
    T07_ConvertVariableFieldsStringRepresentations_01_FourArgs
    T08_ConvertVariableFieldsItemingRepresentations_01_FiveArgs_withobject
    T09_ConvertDoubleQUotes
    T10_ConvertSingleQUotes
    T11_ConvertSmartSingleQUotes
    T12_ConvertSmartSingleQUotes
    ' T13_ConvertVariableFields_DefaultParamSeparator
    ' T14_ConvertVariableFields_AltParamSeparator

    Debug.Print "Testing completed "
    
End Sub


'@TestMethod("Fmt")
Public Sub T01_ConvertFormatFieldXX0ToVbNullString_01()

 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
   'On Error GoTo TestFail
    'Arrange
    Dim myExpected                      As String
    myExpected = "Hello    Hello"
    
    Dim myTest As String
    myTest = myExpected '"Hello {nl0} {tb0} {nt0} Hello"
    
    Dim myResult As String
    
    'Act
    myResult = Fmt.ConvertFormatFieldXX0ToNoString(myTest)
    
    'Assert
    AssertStrictAreEqual myExpected, myExpected, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub


    '@TestMethod("Fmt")
    Public Sub T02_ConvertFormatFieldXXToXX1_01()

        Dim myExpected                      As String
        Dim myTest                          As String
        Dim myResult                        As String

 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
       'On Error GoTo TestFail

        'Arrange
        myExpected = "Hello {nl1} {tb1} {nt1} Hello"
        myTest = "Hello {nl} {tb} {nt} Hello"

        'Act
        myResult = Fmt.ConvertFormatFieldWithNoCountToCountOfOne(myTest)

        'Assert
        AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T03_GetFieldCount_01_nl3_equals_3()


 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
        Dim myExpected                     As Long
        Dim myResult                         As Long

       'On Error GoTo TestFail

        'Arrange
        myExpected = 3

        'Act
        myResult = Fmt.GetRepeatCountForFormatField("Hello {nl3} Hello", "{nl")

        'Assert
        AssertStrictAreEqual myExpected, myResult, myProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T04_GetFormattingReplaceString_01_nl3()


 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
        Dim myExpected                     As String
        Dim myTest                         As String

       'On Error GoTo TestFail

        myExpected = String$(3, vbCrLf)              'vbCrLf & vbCrLf & vbCrLf
        myTest = Fmt.pvGetFormattingReplaceString("{nl", 3) ' nl is the code for vbcrlf

        'Debug.Print VBA.Len(myExpected), VBA.Len(myTest), VBA.Len(vbCrLf & vbCrLf & vbCrLf), VBA.Len(String$(3, vbCrLf))
        AssertStrictAreEqual myExpected, myTest, myProcedureName
        'TestExit:
        Exit Sub
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T05_GetFormattingReplaceString_02_tb4()


 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
        Dim myExpected                     As String
        Dim myTest                         As String

       'On Error GoTo TestFail

        myExpected = vbTab & vbTab & vbTab & vbTab
        myTest = Fmt.pvGetFormattingReplaceString("{tb", 4)
        'Debug.Print "02", VBA.Len(myExpected), VBA.Len(myTest), VBA.Len(vbTab & vbTab & vbTab)
        AssertStrictAreEqual myExpected, myTest, myProcedureName
        'TestExit:
        Exit Sub
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T06_GetFormattingReplaceString_03_nt2()


 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
    
        Dim myExpected                     As String
        Dim myTest                         As String

       'On Error GoTo TestFail

        'Arrange
        myExpected = String$(2, vbCrLf) & vbTab      'vbCrLf & vbCrLf & vbTab

        'Act
        myTest = Fmt.pvGetFormattingReplaceString("{nt", 2)
        'Debug.Print "03", VBA.Len(myExpected), VBA.Len(myTest), VBA.Len(vbCrLf & vbCrLf & vbTab), VBA.Len(String$(2, vbCrLf) & vbTab)
        'Assert
        AssertStrictAreEqual myExpected, myTest, myProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T07_ConvertVariableFieldsStringRepresentations_01_FourArgs()


 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
        Dim myArray                                 As Variant
        Dim myExpected                              As String
        Dim myTest                                  As String
        Dim myResult                                As String

       'On Error GoTo TestFail

        'Arrange:
        myArray = Array(1, "Hello World", True, 3.134)
        myExpected = "Integer 1: 1, Text: Hello World, Boolean: True, Double: 3.134"
        myTest = "Integer 1: {0}, Text: {1}, Boolean: {2}, Double: {3}"

        'Act
       
        myResult = Fmt.ConvertVariableFieldsToStringRepresentations(myTest, myArray)

        'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T08_ConvertVariableFieldsItemingRepresentations_01_FiveArgs_withobject()


 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
        Dim myArray                                 As Variant
        Dim myExpected                              As String
        Dim myResult                                As String
        Dim myTest                                  As String
       'On Error GoTo TestFail
        
        Dim myColl As Collection
        Set myColl = New Collection
        myColl.Add 10
        myColl.Add "Hello"
        myColl.Add True
        myColl.Add 4.2
        
        'Arrange:
        myArray = Array(1, "Hello World", True, 5.134, myColl)
        myExpected = "Integer 1: 1, Text: Hello World, Boolean: True, Double: 5.134, Object {10,Hello,True,4.2}"   ' should object data be in brackets?
        myTest = "Integer 1: {0}, Text: {1}, Boolean: {2}, Double: {3}, Object {4}"

        'Act
        myResult = Fmt.ConvertVariableFieldsToStringRepresentations(myTest, myArray)

        'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T09_ConvertDoubleQUotes()


 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
    
        Dim myExpected                              As String
        Dim myResult                                As String
        Dim myTest                                  As String
       'On Error GoTo TestFail

        'Arrange:
        myExpected = "Should have double quotes ""Hello World"""
        myTest = "Should have double quotes {dq}{0}{dq}"

        'Act
        myResult = Fmt.Txt(myTest, "Hello World")

        'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T10_ConvertSingleQUotes()

 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
        Dim myExpected                              As String
        Dim myResult                                As String
        Dim myTest                                  As String
       'On Error GoTo TestFail

        'Arrange:
        myExpected = "Should have double quotes 'Hello World'"
        myTest = "Should have double quotes {sq}{0}{sq}"

        'Act
        myResult = Fmt.Txt(myTest, "Hello World")

        'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T11_ConvertSmartSingleQUotes()

 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
        Dim myExpected                              As String
        Dim myResult                                As String
        Dim myTest                                  As String
       'On Error GoTo TestFail

        'Arrange:
        myExpected = "Should have single smart quotes " & Char.twLSmartSQuote & "Hello World" & Char.twRSmartSQuote
        myTest = "Should have single smart quotes {so}{0}{sc}"

        'Act
        myResult = Fmt.Txt(myTest, "Hello World")

        'Assert:
        Assert.Permissive.AreEqual myExpected, myResult, myProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T12_ConvertSmartSingleQUotes()

 #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
        Dim myExpected                              As String
        Dim myResult                                As String
        Dim myTest                                  As String
       'On Error GoTo TestFail

        'Arrange:
        myExpected = "Should have double quotes " & Char.twLSMartDQuote & "Hello World" & Char.twRSmartDQuote
        myTest = "Should have double quotes {do}{0}{dc}"

        'Act
        myResult = Fmt.Txt(myTest, "Hello World")

        'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



    '@TestMethod("Fmt")
    Public Sub T13_ConvertVariableFields_DefaultParamSeparator()
    
    
     #If twinbasic Then
    
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
        
        
    #Else
    
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
        
    
    #End If
    
    Dim myExpected                              As String
    Dim myResult                                As String
    Dim myTest                                  As String
   'On Error GoTo TestFail
    
    'Arrange:
    
    myExpected = "1Hello WorldTrue5.134[6,7,8,9]"
    myTest = "{0}{1}{2}{3}{4}"
    
    'Act
    myResult = Fmt(myTest, 1, "Hello World", True, 5.134, Array(6, 7, 8, 9))
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
    'TestExit:
    Exit Sub
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    ' '@TestMethod("Fmt")
    ' Public Sub T14_ConvertVariableFields_AltParamSeparator()
    
    
    '    Dim myExpected                              As String
    '    Dim myResult                                As String
    '    Dim myTest                                  As String
    '   'On Error GoTo TestFail
    
    '    'Arrange:
    
    '    myExpected = "1;;Hello World;;True;;5.134;;[6,7,8,9]"
    '    myTest = "{0}{1}{2}{3}{4}"
    
    
    '    'Act
    '    myResult = Fmt.Txt(myTest, 1, "Hello World", True, 5.134, Array(6, 7, 8, 9))
    
    '    'Assert:
    '    AssertStrictAreEqual myExpected, myResult  , myProcedureName
    
    '    'TestExit:
    '    Exit Sub
    ' TestFail:
    '    Debug.Print CurrentSOurceFile & ":" & CurrentProcedureName &  "." & CurrentProcedureName & "  raised an error: #" & Err.Number & " - " & Err.Description
    ' End Sub

