VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
Option Explicit

Private Const TODAY                         As String = "\Day05.txt"

Private Type State
    Words                                   As seqC
End Type

Private s                                   As State

Public Sub Execute()
    'testNiceV1
    Part01
    Part02
End Sub

Public Sub testNiceV1()
    Debug.Print HasBadSubstrings("Hello ab", "ab,cd,pq,xy"), True
    Debug.Print HasBadSubstrings("Hello ef", "ab,cd,pq,xy"), False
    Debug.Print HasDoubleChars("Hello World"), True
    Debug.Print HasDoubleChars("Hi World"), False
    Debug.Print HasThreeVowels("Hello World"), True
    Debug.Print HasThreeVowels("zxyyz plugh"), False
End Sub


Public Sub Part01()

    Initialise
    Dim myNiceWords As Long
    myNiceWords = 0
    Dim myWords As Iteritems: Set myWords = Iteritems(s.Words)
    Do
        Dim myWord As String: myWord = myWords.curitem(0)
        If IsNiceV1(myWord) Then
            myNiceWords = myNiceWords + 1
        End If

    Loop While myWords.MoveNext

    fmt.dbg "The answer to Day{0} Part 01 is 238. Found is {1}", VBA.Mid$(TODAY, 5, 2), myNiceWords

End Sub


Public Sub Part02()

    Initialise
    Dim myNiceWords As Long
    Dim myWords As Iteritems: Set myWords = Iteritems(s.Words)
    Do
        Dim myWord As String: myWord = myWords.curitem(0)
        If IsNiceV2(myWord) Then
            myNiceWords = myNiceWords + 1
        End If

    Loop While myWords.MoveNext

    fmt.dbg "The answer to Day{0} Part 02 is 69. Found is {1}", VBA.Mid$(TODAY, 5, 2), myNiceWords

End Sub


Private Sub Initialise()
    Set s.Words = seqC(Filer.GetFileAsArrayOfStrings(AoCRawData & Year & TODAY))
End Sub


Public Function IsNiceV1(ByVal ipWord As String) As Boolean

    IsNiceV1 = False
    
    If HasBadSubstrings(ipWord, "ab,cd,pq,xy") Then Exit Function
    If Not HasDoubleChars(ipWord) Then Exit Function
    If Not HasThreeVowels(ipWord) Then Exit Function
    
    IsNiceV1 = True

End Function


Public Function HasBadSubstrings(ByRef ipWord As String, ByRef ipBadSubstrings As String) As Boolean

    Dim myBadSubstrings As Variant: myBadSubstrings = VBA.Split(ipBadSubstrings, ",")
    
    Dim myItem As Variant
    For Each myItem In myBadSubstrings
    
        Dim myBadString As String: myBadString = myItem
        
        If InStr(ipWord, myBadString) > 0 Then

            HasBadSubstrings = True
            Exit Function

        End If

    Next

    HasBadSubstrings = False

End Function


Public Function HasThreeVowels(ByRef ipWord As String) As Boolean

    Dim myVowels As Variant: myVowels = Split("a,e,i,o,u", ",")
    Dim myItem As Variant
    Dim myVowelCount As Long
    For Each myItem In myVowels
        Dim myChar As String: myChar = myItem
        myVowelCount = myVowelCount + VBA.Len(ipWord) - VBA.Len(VBA.Replace(ipWord, myChar, vbNullString))

    Next

    HasThreeVowels = myVowelCount > 2

End Function


Private Function HasDoubleChars(ByVal ipWord As String) As Boolean

    Dim myIndex As Long
    For myIndex = 1 To Len(ipWord) - 1

        If Mid$(ipWord, myIndex, 1) = Mid$(ipWord, myIndex + 1, 1) Then

            HasDoubleChars = True
            Exit Function

        End If

    Next

    HasDoubleChars = False

End Function


Public Function IsNiceV2(ByRef ipWord As String) As Boolean

    IsNiceV2 = False
    If NoRepeatingPairs(ipWord) Then Exit Function
    If NoSpacedRepeats(ipWord) Then Exit Function
    IsNiceV2 = True

End Function


Private Function NoRepeatingPairs(ByRef ipWord As String) As Boolean

    NoRepeatingPairs = True

    Dim myIndex As Long
    For myIndex = 1 To Len(ipWord) - 3

        Dim mySubStr As String
        mySubStr = Mid$(ipWord, myIndex, 2)
        If Len(ipWord) - Len(Replace(ipWord, mySubStr, vbNullString)) > 3 Then

            NoRepeatingPairs = False
            Exit Function

        End If

    Next

End Function


Private Function NoSpacedRepeats(ByRef ipWord As String) As Boolean

    NoSpacedRepeats = True

    Dim myIndex As Long
    For myIndex = 1 To Len(ipWord) - 1

        If (Len(ipWord) - Len(Replace(ipWord, Mid$(ipWord, myIndex, 1), vbNullString))) > 1 Then

            If Mid$(ipWord, myIndex, 1) = Mid$(ipWord, myIndex + 2, 1) Then

                NoSpacedRepeats = False
                Exit Function

            End If

        End If

    Next

End Function
