VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@PredeclaredId
Private Const RAW_DATA_PATH         As String = "C:\Users\slayc\source\repos\VBA\AdventOfCode\2015\Day03RawData.txt"

Private Type State

    RawData                         As Kvp
    
End Type

Private s                           As State


Public Sub Part1()

    Set s.RawData = New Kvp
    s.RawData.AddByIndexAsCharacters Common.GetRawDataByFile(RAW_DATA_PATH)

    Dim myAtLeastOnePresent As Long
    myAtLeastOnePresent = DeliverAtLeastOnePresent(s.RawData)
    
    Debug.Print "The answer to Day 3 Part 1 is 2565. Found is ", myAtLeastOnePresent
    
End Sub


Public Function DeliverAtLeastOnePresent(ByVal ipInstructions As Kvp) As Long

    Dim myJourney As Journey
    Set myJourney = Journey.Make(ipInstructions)
    DeliverAtLeastOnePresent = myJourney.GetVisits.Count

End Function


Public Sub Part2()

    Dim myPresents As Long
    myPresents = GetHousesWithMoreThanOnePresent(s.RawData)
    Debug.Print "The answer to Day 3 Part 2 is 2639. Found is ", myPresents

End Sub

Public Function GetHousesWithMoreThanOnePresent(ByVal ipInstructions As Kvp) As Long

    Dim mySantaInstructions As Kvp: Set mySantaInstructions = New Kvp
    Dim myRobotInstructions As Kvp: Set myRobotInstructions = New Kvp
    
    Dim myItem As Long
    For myItem = 0 To s.RawData.Count - 1 Step 2
    
        mySantaInstructions.AddByIndex ipInstructions.GetItem(myItem)
        myRobotInstructions.AddByIndex ipInstructions.GetItem(myItem + 1)
    
    Next
    
    ' Get Santa and RObot journeys
    Dim mySantaJourney As Journey
    Set mySantaJourney = Journey.Make(mySantaInstructions)
    
    Dim myrobotJourney As Journey
    Set myrobotJourney = Journey.Make(myRobotInstructions)

    mySantaJourney.GetVisits.Cohorts(myrobotJourney.GetVisits).GetItem(Cohort.Cohort_AllAandBOnly).Count
    GetHousesWithMoreThanOnePresent = mySantaJourney.GetVisits.Cohorts(myrobotJourney.GetVisits).GetItem(Cohort.Cohort_AllAandBOnly).Count
    
End Function

