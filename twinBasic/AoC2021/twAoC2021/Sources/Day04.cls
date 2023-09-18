VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Const InputData     As String = "Day04.txt"

Private Type State
    
    DrawnNumbers                    As seq
    Boards                          As seq
    LastScore                       As long
    
End Type

Private s                           As State

Public Sub Execute()
    
    Part01
    Part02
    
End Sub


Private Sub Initialise()
    
    ' SPlit the raw data into into a seq of two items, 
    ' (0))= string of drawn numbers, 
    ' (1) = String describing boards, 
    Dim myData As Variant = VBA.Split(Filer.GetFileAsString(AoC2021Data & InputData), vbCrLf & vbCrLf, 2)
    
     ' Convert drawn numbers as a comma separated string to a seq of Long values
    Set s.DrawnNumbers = Seq.Deb(VBA.Split(myData(0), Char.twComma)).MapIt(mpConvert.Deb(ToLong)) _
    
    ' The string describing boards has single didit numbers which
    ' have been padded to two characters by preceding them with a space character
    ' there are two cases <space><space> and <vbCrLf><space>
    ' the extraneous space needs to be eliminated so we can
    ' split the board numbers using a single space
    Dim myReplaceTwoSpaces As Variant = Array(Char.twSpace & Char.twSpace, Char.twSpace)
    Dim myReplacevbCrLfSpace As Variant = Array(vbCrLf & Char.twSpace, vbCrLf)
    myData(1) = Strs.MultiReplacer(myData(1), Array(myReplaceTwoSpaces, myReplacevbCrLfSpace))
   
    ' Create a seq of Boards from a string of all boards
    ' Substrings representing boards are separate by two vbCrLf
    ' The constructor for a Board takes a String of 5 rows,
    ' each row separated by a single vbCrLf
    ' So for now map the single string into board strings
    Set s.Boards = Seq.Deb(VBA.Split(myData(1), vbCrLf & vbCrLf)).MapIt(mpConstruct(Board))
    
End Sub



Public Sub Part01()
    
    Initialise
    Dim myResult As Long
    myResult = GetFirstWInnerScore
    Fmt.Dbg "The answer to Day 04 part 1 is 45031.  Found is {0}", myResult
    
End Sub

Public Sub Part02()
    
    Initialise
    Dim myResult As Long
    myResult = GetLastWinnerScore
    Fmt.Dbg "The answer to Day 04 part 2 is 2568.  Found is {0}", myResult
    
End Sub


Public Function GetFirstWInnerScore() As Long

    Dim myDrawnNumber As IterItems = IterItems(s.DrawnNumbers)
    Do
   
        Dim myBoard As IterItems = IterItems(s.Boards) 'For Each myBoard In s.Boards
        Do
            
            If myBoard.Item.HasNumber(myDrawnNumber.Item) Then

                If myBoard.Item.haswon Then
                    
                    Return myBoard.Item.WinningScore
                    Exit Function
                
                End If
                
            End If
            
        Loop While myBoard.MoveNext
        
    Loop While myDrawnNumber.MoveNext
                    
    Return -1
    
End Function
 
 
Public Function GetLastWinnerScore() As Long

    ' Track the indexes of the winning boards
    Dim myWinningBoards As Seq = Seq.Deb
 
    Dim myDrawnNumber As IterItems = IterItems(s.DrawnNumbers)
    Do
        Dim myNumber As Long = myDrawnNumber.Item
        Dim myBoards As IterItems = IterItems(s.Boards)
        Do
            Dim myBoard As Board = myBoards.Item
            DoEvents
            If myWinningBoards.LacksItem(myBoard.Index) Then
            
                If myBoard.HasNumber(myNumber) Then

                    If myBoard.HasWon Then
                        
                        myWinningBoards.Add myBoard.Index
                        s.LastScore = myBoard.WinningScore
                        
                    End If
                    
                End If
            
            End If
            
        Loop While myBoards.MoveNext
        
    Loop While myDrawnNumber.MoveNext
    
    Return s.LastScore
  
End Function


