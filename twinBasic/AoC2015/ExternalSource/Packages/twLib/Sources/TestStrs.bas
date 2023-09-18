Attribute VB_Name = "TestStrs"
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

Public Sub StrsTests()

#If twinbasic Then
    Debug.Print CurrentProcedureName ; vbTab, vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If
    
    T01a_DedupDefaultSpaceChar
    T01b_DedupCharsInParamArray
    T01c_DedupCharsInArray
    T01d_DedupCharsInString
    T01e_DedupCharsInEnumerable
    
    T02a_TrimmerDefaultChars
    T02b_TrimmerCharsInParamArray
    T02c_TrimmerCharsInArray
    T02d_TrimmerCharsInString
    T02e_TrimmerCharsInEnumerable
    
    T03a_PadLeftDefaultChar
    T03b_PadLeftWithtwHash
    T03c_PadLeftWithtwHashtwPlus
    
    T04a_PadRightDefaultChar
    T04b_PadRightWithtwHash
    T04c_PadRightWithString
    
    T05a_FreqForSingleCharacters
    T05b_FreqSubString
    
    T06a_ToSubStrLystDefaultSeperatorAndTrimChars
    T06b_ToSubStrLystExplicitSeperatorAndDefaultTrimChars
    T06c_ToSubStrLystDefaultSeperatorAndExplicitTrimChars
    T06d_ToSubStrLystExplicitSeperatorAndExplicitTrimChars
    
    T07a_RepeatReplacerDefaults
    T07b_RepeatReplacerExplicitFindChar
    T07c_RepeatReplacerExplicitFindCharExplicitReplaceChar
    
    T08a_MultiReplacerDefaultFindCharDefaultReplaceChar
    T08b_MultiReplacerExplicitFindCharExplicitReplaceChar
    
    T09a_ToCharLystvbNullString
    T09b_ToCharLystEmptyString
    T09c_ToCharLystString
    
    Debug.Print "Testing completed"
    
End Sub

'@TestMethod("Dedup")
Private Sub T01a_DedupDefaultSpaceChar()

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
    myExpected = " Hello Worldee "

     Dim myResult As String

    'Act:
    myResult = Strs.Dedup("     Hello Worldee      ")

    Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub

    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
        
End Sub

'@TestMethod("Dedup")
Private Sub T01b_DedupCharsInParamArray()

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
    myExpected = " Hello Worlde "
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.Dedup("     Hellooo Worrldee      ", "e", "o", "r", " ")

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
End Sub

'@TestMethod("Dedup")
Private Sub T01c_DedupCharsInArray()

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
        myExpected = " Hello Worlde "
        
        
        Dim myResult As String
        
        'Act:
        myResult = Strs.Dedup("     Hellooo Worrldee      ", Array("e", "o", "r", " "))

        'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
End Sub

'@TestMethod("Dedup")
Private Sub T01d_DedupCharsInString()

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
    myExpected = " Hello Worlde "
    
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.Dedup("     Hellooo Worrldee      ", "eor ")

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
        
End Sub

'@TestMethod("Dedup")
Private Sub T01e_DedupCharsInEnumerable()

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
        myExpected = " Hello Worlde "
        
        
        Dim myResult As String
        Dim myList As Seq = Seq.Deb("eor ")
        'Act:
        myResult = Strs.Dedup("     Hellooo Worrldee      ", "eor ")

        'Assert:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
End Sub

'@TestMethod("Trimmer")
Private Sub T02a_TrimmerDefaultChars()

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
    
    'Act:
    myResult = Strs.Trimmer("   Hello World      ")

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
        
End Sub

'@TestMethod("Trimmer")
Private Sub T02b_TrimmerCharsInParamArray()

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
    
    'Act:
    myResult = Strs.Trimmer("   ;;;,;,;Hello World ;,; ;; ,", Char.twSpace, Char.twSemiColon, Char.twComma, Char.twPeriod)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
        
End Sub

'@TestMethod("Trimmer")
Private Sub T02c_TrimmerCharsInArray()

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
    
    'Act:
    myResult = Strs.Trimmer("   ;;;,;,;Hello World ;,; ;; ,", Array(" ", ";", ","))

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Trimmer")
Private Sub T02d_TrimmerCharsInString()

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
    
    'Act:
    myResult = Strs.Trimmer("   ;;;,;,;Hello World ;,; ;; ,", " ;,")

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Trimmer")
Private Sub T02e_TrimmerCharsInEnumerable()

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
    Dim myList As Seq = Seq.Deb(" ;,")
    'Act:
    myResult = Strs.Trimmer("   ;;;,;,;Hello World ;,; ;; ,", myList)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("PadLeft")
Private Sub T03a_PadLeftDefaultChar()

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
    myExpected = "     Hello"
    
    Dim myResult As String

    'Act:
    myResult = Strs.PadLeft("Hello", 10)

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub

TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("PadLeft")
Private Sub T03b_PadLeftWithtwHash()

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
    myExpected = "#####Hello"
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.PadLeft("Hello", 10, Char.twHash)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("PadLeft")
Private Sub T03c_PadLeftWithtwHashtwPlus()

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
    myExpected = "#+#+#+#+#+Hello"
    
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.PadLeft("Hello", 10, Char.twHash & Char.twPlus)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("PadRight")
Private Sub T04a_PadRightDefaultChar()

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
    myExpected = "Hello     "
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.PadRight("Hello", 10)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("PadRight")
Private Sub T04b_PadRightWithtwHash()

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
    myExpected = "Hello#####"
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.PadRight("Hello", 10, Char.twHash)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("PadRight")
Private Sub T04c_PadRightWithString()

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
    myExpected = "Hello#+#+#+#+#+"
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.PadRight("Hello", 10, Char.twHash & Char.twPlus)
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Freq")
Private Sub T05a_FreqForSingleCharacters()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 3
    
    Dim myResult As Long
    
    'Act:
    myResult = Strs.CountOf("Hello World", "l")
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Freq")
Private Sub T05b_FreqSubString()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 1
    
    Dim myResult As Long
    
    'Act:
    myResult = Strs.CountOf("Hello World", "ll")
    
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("SubStrLyst")
Private Sub T06a_ToSubStrLystDefaultSeperatorAndTrimChars()

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
        myExpected = Array("Hello", "its", "a", "nice", "day")
        
        
        Dim myResult As Variant
       
        'Act:
        myResult = Strs.ToSubStr(" Hello ,its  , a, nice, day").ToArray
       
        'Assert:
        AssertStrictSequenceEquals myExpected, myResult, myProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
End Sub

'@TestMethod("SubStrLyst")
Private Sub T06b_ToSubStrLystExplicitSeperatorAndDefaultTrimChars()

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
    myExpected = Array("Hello", "its", "a", "nice", "day")

    Dim myResult As Variant

    'Act:
    myResult = Strs.ToSubStr(" Hello ,its  , a, nice, day", ",").ToArray

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub

TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("SubStrLyst")
Private Sub T06c_ToSubStrLystDefaultSeperatorAndExplicitTrimChars()

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
    myExpected = Array("Hello", "its", "a", "nice", "day")
    
    Dim myResult As Variant
   
    'Act:
     myResult = Strs.ToSubStr(" ,Hello ,its  , a, nice, day", ipTrimChars:=Seq.Deb(Strs.WhiteSpace & Char.twComma).ToArray).ToArray
   
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("SubStrLyst")
Private Sub T06d_ToSubStrLystExplicitSeperatorAndExplicitTrimChars()

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
    myExpected = Array("Hello", "its", "a", "nice", "day")
    
    Dim myResult As Variant
   
    'Act:
     myResult = Strs.ToSubStr(" ;;Hello ;its  ; a; nice; day", Char.twSemiColon, Seq.Deb(Strs.WhiteSpace & Char.twSemiColon).ToArray).ToArray
   
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("RepeatReplacer")
Private Sub T07a_RepeatReplacerDefaults()

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
    myExpected = "HelloWorld"
    
    Dim myResult As Variant
   
    'Act:
     myResult = Strs.RepeatReplacer("  Hel  lo Wo rld       ")
   
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("RepeatReplacer")
Private Sub T07b_RepeatReplacerExplicitFindChar()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

On Error GoTo TestFail

'Arrange:
Dim myExpected  As String
myExpected = "HelloWorld"

Dim myResult As Variant

'Act:
    myResult = Strs.RepeatReplacer("###Hel##lo#Wo##rld#####", Char.twHash)

'Assert:
AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("RepeatReplacer")
Private Sub T07c_RepeatReplacerExplicitFindCharExplicitReplaceChar()

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
    myExpected = "@@@Hel@@lo@Wo@@rld@@@@@"
    
    Dim myResult As Variant
   
    'Act:
     myResult = Strs.RepeatReplacer("###Hel##lo#Wo##rld#####", Char.twHash, Char.twAmp)
   
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("RepeatReplacer")
Private Sub T08a_MultiReplacerDefaultFindCharDefaultReplaceChar()

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
    myExpected = "HelloWorld"
    
    Dim myResult As Variant
   
    'Act:
     myResult = Strs.RepeatReplacer("   Hel  lo Wo  rld     ")
   
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("RepeatReplacer")
Private Sub T08b_MultiReplacerExplicitFindCharExplicitReplaceChar()

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
    myExpected = "HelloWorld"
    
    Dim myResult As String
   
    'Act:
     myResult = Strs.MultiReplacer("##@Hel@@lo#Wo#@rld@@@@@", Array(Array(Char.twHash, vbNullString), Array(Char.twAmp, vbNullString)))
   
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Strs")
Private Sub T09a_ToUnicodeBytesList()
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected()  As Byte
    Dim myTmp As Variant
    myTmp = Array(32, 0, 32, 0, 32, 0, 72, 0, 101, 0, 108, 0, 32, 0, 32, 0, 108, 0, 111, 0, 32, 0, 87, 0, 111, 0, 32, 0, 32, 0, 114, 0, 108, 0, 100, 0, 32, 0, 32, 0, 32, 0, 32, 0, 32, 0)
    ReDim myExpected(LBound(myTmp) To UBound(myTmp))
    Dim myIndex As Long
    For myIndex = LBound(myTmp) To UBound(myTmp)
        myExpected(myIndex) = CByte(myTmp(myIndex))
    Next
    
    Dim myResult As Variant
   
    'Act:
     myResult = Strs.ToUnicodeBytes("   Hel  lo Wo  rld     ").ToArray

    'Assert:
    
    Dim mytest As Long
  
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("ToCharLyst")
Private Sub T09a_ToCharLystvbNullString()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant = Array(" ")
    myExpected(0) = vbNullString
    
    Dim myResult As Variant
   
    'Act:
     myResult = Seq.Deb(vbNullString).ToArray
   
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("ToCharLyst")
Private Sub T09b_ToCharLystEmptyString()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If
    
   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant = Array("")
    
    Dim myResult As Variant
   
    'Act:
    myResult = Seq.Deb("").ToArray
   
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("ToCharLyst")
Private Sub T09c_ToCharLystString()

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
    myExpected = Array("H", "e", "l", "l", "o")
    
    
    Dim myResult As Variant
   
    'Act:
     myResult = Seq.Deb("Hello").ToArray
   
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Strs")
Private Sub T10a_BinToLong()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

   'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long = 42
   
    
    
    Dim myResult As Variant
   
    'Act:
     myResult = Strs.BinToLong("101010")
   
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub