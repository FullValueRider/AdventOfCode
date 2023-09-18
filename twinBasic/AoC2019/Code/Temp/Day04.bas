Attribute VB_Name = "Day04"
Option Explicit
'@Folder("AdventOfCode")

Public Sub Day4Part1()

    Dim myPassword As Long
    For myPassword = 168630 To 718098
        
        Dim myPasswordStr As String
        myPasswordStr = CStr(myPassword)
        
        Dim myCounter As Long
        If DigitsSameOrBiggerFromLeftToRight(myPasswordStr) And HoldsAPair(myPasswordStr) Then myCounter = myCounter + 1
        
    Next
    
    Debug.Print "The answer to Day 04 Part 1 should be should be 1686 ", myCounter
End Sub

Public Sub Day4Part2()
    
    Dim myPassword As Long
    For myPassword = 168630 To 718098
    
        Dim myPasswordStr As String
        myPasswordStr = CStr(myPassword)
        
        Dim myCounter As Long
        If DigitsSameOrBiggerFromLeftToRight(myPasswordStr) And HoldsPairOnly(myPasswordStr) Then myCounter = myCounter + 1
        
    Next
    
    Debug.Print "The answer to Day 04 Part 1 should be should be 1145 ", myCounter
    
End Sub


Public Function DigitsSameOrBiggerFromLeftToRight(ByVal ipNumber As String) As Boolean

    Dim myIndex As Long
    
    For myIndex = Len(ipNumber) To 2 Step -1
    
        If Mid$(ipNumber, myIndex, 1) < Mid$(ipNumber, myIndex - 1, 1) Then
        
            DigitsSameOrBiggerFromLeftToRight = False
            Exit Function
            
        End If
    
    Next
    
    DigitsSameOrBiggerFromLeftToRight = True
    
End Function

Public Function HoldsAPair(ByVal ipNumber As String) As Boolean

    Dim myIndex As Long
    For myIndex = 0 To 9
    
        Dim myPair As String
        myPair = String$(2, CStr(myIndex))
        
        If InStr(ipNumber, myPair) > 0 Then
            
            HoldsAPair = True
            Exit Function

        End If
        
    Next
    
    HoldsAPair = False

End Function

Public Function HoldsPairOnly(ByVal ipNumber As String) As Boolean
        
    HoldsPairOnly = False
    
    Dim myDigits As Variant
    myDigits = Split("0,1,2,3,4,5,6,7,8,9", ",")
    
    Dim myNumberLen As Long
    myNumberLen = Len(ipNumber)
    
    Dim myDigit As Variant
    For Each myDigit In myDigits
    
        If myNumberLen - Len(Replace(ipNumber, myDigit, vbNullString)) = 2 Then
        
            If InStr(ipNumber, String$(2, myDigit)) > 0 Then
            
                HoldsPairOnly = True
                Exit Function
            
            End If
        
        End If
        
    Next
    
End Function


