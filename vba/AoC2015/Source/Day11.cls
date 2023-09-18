VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
 '@PredeclaredId
 '@Exposed
 '@Folder("Days")
 Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C


Private Const TODAY             As String = "\Day11.txt"

Private Type State
    Data                            As String
End Type

Private s                           As State


Public Sub Execute()
    
        Part01
        Part02
        
End Sub

Private Sub Part01()

    Initialise
    
    Dim myString As String: myString = "hepxcrrq"
    
    Do
    
        myString = strs.Inc(myString, chars.Lower)
 
        If IsLegalPassWord(myString) Then
            Exit Do
        End If
       
    Loop
    
    Dim myResult As String: myResult = myString
    
    fmt.dbg "The answer to Day {0} part 1 is {1}.  Found is {2}", VBA.Mid$(TODAY, 5, 2), "hepxxyzz", myResult
    
End Sub

Private Sub Part02()

    Initialise
    
    Dim myString As String: myString = s.Data
    Do
        myString = strs.Inc(myString, chars.Lower)
        If IsLegalPassWord(myString) Then
            Exit Do
        End If
    Loop
    
    Do
        myString = strs.Inc(myString, chars.Lower)
        If IsLegalPassWord(myString) Then
            Exit Do
        End If
    Loop
  
    Dim myResult As String: myResult = myString
    
    fmt.dbg "The answer to Day {0} part 2 is {1}.  Found is {2}", VBA.Mid$(TODAY, 5, 2), "heqaabcc", myResult

End Sub



' Process the incoming data to a suitable format
Private Sub Initialise()

    s.Data = Filer.GetFileAsString(AoCRawData & Year & TODAY)
        
End Sub


Private Function IsLegalPassWord(ByRef ipString As String) As Boolean

    IsLegalPassWord = False
    If Not HasTwoDifferentNonOverlappingPairs(ipString) Then Exit Function
    If HasIllegalLetters(ipString, "iol") Then Exit Function
    If Not HasThreeIncreasingChars(ipString) Then Exit Function
    IsLegalPassWord = True
    
End Function

Public Function HasTwoDifferentNonOverlappingPairs(ByRef ipString As String) As Boolean

    HasTwoDifferentNonOverlappingPairs = True
    
    Dim myChars As Iteritems: Set myChars = Iteritems(ipString)
    Dim myS As seqC: Set myS = seqC.Deb
    Do While myChars.MoveNext
        'Debug.Print myChars.Item(-1), myChars.Item(0)
        If myChars.curitem(-1) = myChars.curitem(0) Then
            Dim myPair As String: myPair = myChars.curitem(-1) & myChars.curitem(0)
            If myS.LacksItem(myPair) Then
                myS.Add myPair
                If myS.Count = 2 Then
                    Exit Function
                End If
            End If
            myChars.MoveNext
        End If
            
    Loop
    
    HasTwoDifferentNonOverlappingPairs = False
    
End Function


Public Function HasIllegalLetters(ByRef ipString As String, ByRef ipIllegal As String) As Boolean

    HasIllegalLetters = True
    Dim myIndex As Long
    For myIndex = 1 To VBA.Len(ipIllegal)
        If InStr(ipString, VBA.Mid$(ipIllegal, myIndex, 1)) > 0 Then Exit Function
    Next
    HasIllegalLetters = False
    
End Function


Public Function HasThreeIncreasingChars(ByVal ipString As String) As Boolean

    HasThreeIncreasingChars = True
   ' Debug.Print ipString
    Dim myChars As Iteritems: Set myChars = Iteritems(ipString).FTS(3)
    Do
       ' Debug.Print mychars.Item(-2), mychars.Item(-1), mychars.Item(0)
        If VBA.Asc(myChars.curitem(-2)) = VBA.Asc(myChars.curitem(-1)) - 1 Then
            If VBA.Asc(myChars.curitem(-1)) = VBA.Asc(myChars.curitem(0)) - 1 Then
                Exit Function
            End If
        End If
    Loop While myChars.MoveNext
    
    HasThreeIncreasingChars = False
    
End Function
