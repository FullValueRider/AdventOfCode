Attribute VB_Name = "Day06"
Option Explicit

'@Folder("AdventOfCode")

Private Const HOST                                      As Long = 0
Private Const Moon                                      As Long = 1
Private Const DAY06_INPUT_PATH_AND_NAME                 As String = "C:\Users\slayc\source\repos\VBA\AdventOfCode\2019\Day06Input.txt"


Private Type State

    ListOfMoonsVsHost                                    As KvpOD ' Key is name of body, item is name of host

End Type

Private s                                           As State

Public Sub Day6CodePart1()

    PopulateKvpOfMoonsVsHost GetFileByLines(DAY06_INPUT_PATH_AND_NAME)
    
    Debug.Print Fmt("The answer to Day 6 Part 01 should be 142497: {0}", CalculateTotalOrbits)

End Sub


Public Sub Day6Part2()

    PopulateKvpOfMoonsVsHost GetFileByLines(DAY06_INPUT_PATH_AND_NAME)

    Dim SantaOrbit As KvpOD
    Set SantaOrbit = GetPathToCom("SAN")

    Dim myOrbit As KvpOD
    Set myOrbit = GetPathToCom("YOU")

    Dim myIntersect As KvpOD
    Set myIntersect = SantaOrbit.Cohorts(myOrbit)

    Dim myMovement As Long
    myMovement = SantaOrbit.Count + myOrbit.Count - (2 * myIntersect.Item(KeyCohort_Common).Count) - 2
    Debug.Print Fmt("The answer to Day 06 Part 2 should be 301: {0}", myMovement)

End Sub


Public Sub PopulateKvpOfMoonsVsHost(ByVal ipListOfBodyPairs As KvpOD)

    ' We are supplied with a list in the form Host)Moon
    ' we need to build a list of Moon)Host because
    ' moons can only have one host.
    ' We also need to remember that 'COM' will not appear as a key
    ' in the myMoons Kvp.  This absense of a the Host as a key in
    ' myMoons means we have found 'COM'
    
    Set s.ListOfMoonsVsHost = New KvpOD
    
    Dim myKV As KVPair
    For Each myKV In ipListOfBodyPairs
    
        Dim mySystem As Variant
        mySystem = Split(myKV.Value, ")")
       
        s.ListOfMoonsVsHost.AddByKey mySystem(Moon), mySystem(HOST)
        
     Next
    
End Sub


Public Function GetPathToCom(ByVal ipStartName As String) As KvpOD

    Dim myPath As KvpOD: Set myPath = New KvpOD
    Dim myCurrentBody As String: myCurrentBody = ipStartName
    
    Do While s.ListOfMoonsVsHost.HoldsKey(myCurrentBody)

        Dim myNextBody As String
        myNextBody = s.ListOfMoonsVsHost.Item(myCurrentBody)
        
        myPath.AddByKey myCurrentBody, myNextBody
        myCurrentBody = myNextBody

    Loop

    Set GetPathToCom = myPath

End Function


'@Ignore FunctionReturnValueNotUsed
Public Function CalculateTotalOrbits() As Long

    Dim myKV As KVPair
    For Each myKV In s.ListOfMoonsVsHost

        Dim myCount As Long
        myCount = myCount + GetPathToCom(myKV.Value).Count + 1
        
    Next

    CalculateTotalOrbits = myCount

End Function
