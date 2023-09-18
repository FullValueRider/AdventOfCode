Attribute VB_Name = "Day4"
Option Explicit
'@Folder("AdventOfCode")

Public Function IsPassWordV1(ByVal ipNumber As Long) As Boolean

Dim myArray(1 To 6)                     As Long
'@Ignore VariableNotUsed, VariableNotAssigned
Dim myString                            As String
Dim myIndex                             As Long
Dim myLower                             As Long
Dim myUpper                             As Long
Dim PairFound                           As Boolean

    
    myLower = LBound(myArray)
    myUpper = UBound(myArray)
    
    For myIndex = myLower To myUpper
    
        myArray(myIndex) = CLng(Mid$(CStr(ipNumber), myIndex, 1))
        
    Next
    
    IsPassWordV1 = False
    PairFound = False
    For myIndex = myUpper To myLower + 1 Step -1
    
        If myArray(myIndex) < myArray(myIndex - 1) Then Exit Function
        
        If myArray(myIndex) = myArray(myIndex - 1) Then
        
            PairFound = True
            
        End If
        
    Next
    
    If Not PairFound Then Exit Function
    
    IsPassWordV1 = True

End Function

Public Function IsPassWordV2(ByVal ipNumber As Long) As Boolean

Dim myArray(1 To 6)                     As Long
'@Ignore VariableNotUsed, VariableNotAssigned
Dim myString                            As String
Dim myIndex                             As Long
Dim myLower                             As Long
Dim myUpper                             As Long
'@Ignore VariableNotUsed
Dim PairFound                           As Boolean

    
    myLower = LBound(myArray)
    myUpper = UBound(myArray)
    
    For myIndex = myLower To myUpper
    
        myArray(myIndex) = CLng(Mid$(CStr(ipNumber), myIndex, 1))
        
    Next
    
    IsPassWordV2 = False
    
    For myIndex = myUpper To myLower + 1 Step -1
    
        If myArray(myIndex) < myArray(myIndex - 1) Then Exit Function
        
    Next

    ' Password contains same or increasing numbers
    
    If Not HoldsPair(CStr(ipNumber)) Then Exit Function
    
    IsPassWordV2 = True

End Function

Public Function HoldsPair(ByVal ipNumber As String) As Boolean

Dim myIndex                             As Long

    HoldsPair = False
    
    For myIndex = 0 To 9
    
        If InStr(ipNumber, String$(2, CStr(myIndex))) > 0 Then
        
            If InStr(ipNumber, String$(3, CStr(myIndex))) = 0 Then
            
                HoldsPair = True
                Exit Function
                
            End If
            
        End If
        
    Next
    
End Function


Public Sub GetLegalPasswords()

Dim myCounter                           As Long
Dim myIndex                             As Long

    myCounter = 0
    
    For myIndex = 168630 To 718098
    
        If IsPassWordV2(myIndex) Then myCounter = myCounter + 1
        
    Next
    
    Debug.Print myCounter
    
End Sub
