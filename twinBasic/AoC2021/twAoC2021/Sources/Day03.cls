VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

/* 
The most common bit is determin by countting the 1's and doubling 
the value then subtracting it from the count and taking the 
sign of the result.
A negative value means that ones were more than half the count
*/
Private Enum CommonBit
    One = -1    ' Ones are most common
    Equal = 0   ' Equal 1 and 0
    Zero = 1    ' Zeros are most common
End Enum
Const InputData     As String = "Day03.txt"

Private Type State

    Data As Seq
  
End Type

Private s                       As State


Public Sub Execute()
    
    Part01
    Part02
    
End Sub

Public Sub Part01()
    
    If s.Data Is Nothing Then
       
        Dim myStrings As Variant = Filer.GetFileAsArrayOfStrings(AoC2021Data & InputData, vbCrLf)
        Set s.Data = Seq.Deb.AddRange(myStrings)


    End If

    Dim myGamma As Long = 0
    Dim myEpsilon As Long = 0
   
    Dim myI As IterNum = IterNum(1, VBA.Len(s.Data.First))
    Do
        Dim myBits As Variant
        
        Dim myMostCommonBit As CommonBit = VBA.Sgn(s.Data.Count - ((s.Data.CountIt(cmpEQ("1", myI.Item)) << 1)))

        Select Case myMostCommonBit
            Case One:               myBits = Array(1, 0)
            Case Equal:             myBits = Array(1, 0)
            Case Zero:              myBits = Array(0, 1)
        End Select
       
        myGamma <<= 1
        myEpsilon <<= 1
        myGamma += myBits(0)
        myEpsilon += myBits(1)
        
    Loop While myI.MoveNext
     
    Fmt.Dbg "The answer to Day 03 part 1 is 845186.  Found is {0}", myGamma * myEpsilon
    
End Sub

Public Sub Part02()
    
    Dim myStrings As Variant = Filer.GetFileAsArrayOfStrings(AoC2021Data & InputData, vbCrLf)
    Set s.Data = Seq.Deb.AddRange(myStrings)

    Dim myFoundBit As String
    Dim myOxygen As Seq = s.Data.Clone
    'Dim myMostCommonBit As MostCommonBit

    Dim myI As IIterator = IterNum(1, VBA.Len(myOxygen.First))
    Do

        Dim myMostCommonBit As CommonBit = VBA.Sgn(myOxygen.Count - ((myOxygen.CountIt(cmpEQ("1", myI.Item)) << 1)))

        Select Case myMostCommonBit
            Case One:               myFoundBit = "1"
            Case Equal:             myFoundBit = "1"
            Case Zero:              myFoundBit = "0"
        End Select
        Set myOxygen = myOxygen.FilterIt(cmpEQ.Deb(myFoundBit, myI.Item))
        If myOxygen.Count = 1 Then Exit Do
     
    Loop While myI.MoveNext
    
    Dim myCarbonDioxide As Seq = s.Data.Clone
    Set myI = IterNum(1, VBA.Len(myCarbonDioxide.First))
    Do
        myMostCommonBit = VBA.Sgn(myCarbonDioxide.Count - ((myCarbonDioxide.CountIt(cmpEQ("1", myI.Item)) << 1)))
        Select Case myMostCommonBit
            Case One:       myFoundBit = "0"
            Case Equal:     myFoundBit = "0"
            Case Zero:      myFoundBit = "1"
        End Select
        Set myCarbonDioxide = myCarbonDioxide.FilterIt(cmpEQ.Deb(myFoundBit, myI.Item))
        If myCarbonDioxide.Count = 1 Then Exit Do
     
    Loop While myI.MoveNext
    
    Dim myO As Long = Strs.BinToLong(myOxygen.First)
    Dim myC As Long = Strs.BinToLong(myCarbonDioxide.First)
    Dim myResult As Long = myO * myC
    Fmt.Dbg "The answer to Day 03 Part 2 is 4636702. Found {0}", myResult
    
    End Sub