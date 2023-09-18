VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Days")
Option Explicit

Private Const TODAY                 As String = "\Day03.txt"


Private Type State
    Moves                           As SeqA
End Type

Private s                           As State

Public Sub Execute()
    Part01
    Part02
End Sub

Private Sub Initialise()
    Set s.Moves = SeqA(Filer.GetFileAsString(AllDays.AoCRawData & AllDays.Year & TODAY))
End Sub

Private Sub Part01()
    
    Initialise
    
    Dim myJourney As Journey
    Set myJourney = Journey(s.Moves)
    
    Dim myResult As Long
    myResult = myJourney.Visits.Count
    
    fmt.Dbg "The answer to Day {0} Part 01 is 2572.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult
    
End Sub


Private Sub Part02()

    Initialise

    Dim mySantaJourney As Journey
    Set mySantaJourney = Journey(s.Moves.Slice(1, ipStep:=2))
    
    Dim myRobotJourney As Journey
    Set myRobotJourney = Journey(s.Moves.Slice(2, ipStep:=2))
    
    Dim myResult As Long
    myResult = seqC(mySantaJourney.Visits.Keys).SetOf(e_Setof.m_Unique, seqC(myRobotJourney.Visits.Keys)).Count
   
    fmt.Dbg "The answer to Day {0} Part 02 is 2631.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult
    
End Sub
