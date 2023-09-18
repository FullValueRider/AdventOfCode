VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
Option Explicit

Private Type State

    BadSubStrings                           As Variant
    Vowels                                  As Variant
    Words                                   As Hkvp

End Type

Private s                                   As State

Private Sub Class_Initialize()

    s.BadSubStrings = Split("ab,cd,pq,xy", ",")
    s.Vowels = Split("a,e,i,o,u", ",")
    Set s.Words = Common.GetRawDataByLine("C:\Users\slayc\source\repos\VBA\AdventOfCode\2015\Day05RawData.txt")

End Sub

Public Sub Part1()

    Dim myWord As Variant
    For Each myWord In s.Words

        If IsNiceV1(myWord) Then

            Dim myNiceWords As Long
            myNiceWords = myNiceWords + 1

        End If

    Next

    Debug.Print Layout.Fmt("The answer for Day05 Part 1 should be 258. Found is {0}", myNiceWords)

End Sub


Public Sub Part2()

    Dim myWord As Variant
    For Each myWord In s.Words

        If IsNiceV2(myWord) Then

            Dim myNiceWords As Long
            myNiceWords = myNiceWords + 1

        End If

    Next

    Debug.Print Layout.Fmt("The answer for Day05 Part 2 is xxxx. Found is {0}", myNiceWords)

End Sub


Public Function IsNiceV1(ByVal ipWord As String) As Boolean

    IsNiceV1 = False
    If HasBadSubstrings(ipWord) Then Exit Function
    If HasNoDoubleChars(ipWord) Then Exit Function
    If HasLessThanThreeVowels(ipWord) Then Exit Function
    IsNiceV1 = True

End Function


Public Function HasBadSubstrings(ByVal ipWord As String) As Boolean

    Dim myItem As Variant
    For Each myItem In s.BadSubStrings

        If InStr(ipWord, myItem) > 0 Then

            HasBadSubstrings = True
            Exit Function

        End If

    Next

    HasBadSubstrings = False

End Function


Public Function HasLessThanThreeVowels(ByVal ipWord As String) As Boolean

    Dim myItem As Variant
    For Each myItem In s.Vowels

        Dim myVowelCount As Long
        myVowelCount = myVowelCount + Len(ipWord) - Len(Replace(ipWord, myItem, vbNullString))

    Next

    HasLessThanThreeVowels = IIf(myVowelCount < 3, True, False)

End Function


Private Function HasNoDoubleChars(ByVal ipWord As String) As Boolean

    Dim myIndex As Long
    For myIndex = 1 To Len(ipWord) - 1

        If Mid$(ipWord, myIndex, 1) = Mid$(ipWord, myIndex + 1, 1) Then

            HasNoDoubleChars = False
            Exit Function

        End If

    Next

    HasNoDoubleChars = True

End Function


Public Function IsNiceV2(ByVal ipWord As String) As Boolean

    IsNiceV2 = False
    If NoRepeatingPairs(ipWord) Then Exit Function
    If NoSpacedRepeats(ipWord) Then Exit Function
    IsNiceV2 = True

End Function


Private Function NoRepeatingPairs(ByVal ipWord As String) As Boolean

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


Private Function NoSpacedRepeats(ByVal ipWord As String) As Boolean

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
