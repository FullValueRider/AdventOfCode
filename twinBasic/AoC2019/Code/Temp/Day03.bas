Attribute VB_Name = "Day03"
'@IgnoreModule
' The learning from Day 3 is that we should include comments that describe the structre of any Kvp present.
' That this should be done at the top of the module so a split screen can be used to view definitions whilst coding.

    'KvpWires:                  Index:Long by Links:String->"U32,D56,R79,D354"
    'KvpWireAsLinks             Index:Long by Link:String->"U32"
    'KvpWireAsPath              Point:String->"(x,y)" by Kvp-> Index:Long by TotalSteps:Long
    'KvpWirePathIntersections   Point:String->"(x,y)" by kvp-> Index:long by XYSteps:String->"(XTotalSteps,YTotalSteps)"
    
Option Explicit

Const WIRE_1                        As Long = 0
Const WIRE_2                        As Long = 1

Const DAY03_INPUT_PATH_AND_NAME         As String = "C:\Users\slayc\source\repos\VBA\AdventOfCode\2019\Day03Input.txt"



Public Sub Part1()

    Dim myMinManhatten As Long
    myMinManhatten = GetMinManhattenDistance(Day03Common)
    
    Debug.Print "The answer for Day 03 Part 1 should be 1983: "; myMinManhatten
    
End Sub


Public Sub Part2()

    Dim myMinimumSteps As Long
    myMinimumSteps = GetIntersectionWithMinimumSteps(Day03Common)
    
    Debug.Print "The answer for Day 03 Part 2 should be 107754: "; myMinimumSteps
    
End Sub

Public Function Day03Common() As KvpWirePathIntersections

    Dim myWires As KvpWires: Set myWires = New KvpWires
    Set myWires.Kvp = GetFileByLines(DAY03_INPUT_PATH_AND_NAME)
    
    Dim myWire1AsLinks As KvpWireAsLinks: Set myWire1AsLinks = New KvpWireAsLinks
'    Dim myVar As Variant
'    myVar = myWires.Item(WIRE_1)
'    Debug.Print myWires.Item(WIRE_1)
'    myVar =
    myWire1AsLinks.AddByIndexFromArray Split(myWires.Item(WIRE_1), ",")
    
    Dim myWire2AsLinks As KvpWireAsLinks: Set myWire2AsLinks = New KvpWireAsLinks
    Debug.Print myWires.Item(WIRE_2)
    myWire2AsLinks.AddByIndexFromArray Split(myWires.Item(WIRE_2), ",")
    
    Dim myWire1Path As KvpWireAsPath
    Set myWire1Path = GetWirePath(myWire1AsLinks)
    
    Dim myWire2Path As KvpWireAsPath
    Set myWire2Path = GetWirePath(myWire2AsLinks)
    
    Dim myIntersections As KvpWirePathIntersections
    Set Day03Common = GetIntersectionsOfWires(myWire1Path, myWire2Path)
    
End Function

Public Function GetWirePath(ByRef ipWireLinks As KvpWireAsLinks) As KvpWireAsPath
    


    Dim myWirePath As KvpWireAsPath: Set myWirePath = New KvpWireAsPath
    
    
    Dim myItem As Variant
    'Stop
    For Each myItem In ipWireLinks
        Debug.Print myItem
        DoEvents
        Dim myDirection As String
        myDirection = VBA.Left$(myItem.Value, 1)
        
        Dim myLinkSteps As Long
        myLinkSteps = CLng(Mid$(myItem.Value, 2))

        '@Ignore VariableNotUsed, UndeclaredVariable
        Dim myStep As Long
        For myStep = 1 To myLinkSteps
        
            DoEvents
            Dim myX As Long
            Dim myY As Long
            Select Case myDirection

                Case "U": myY = myY + 1
                Case "D": myY = myY - 1
                Case "R": myX = myX + 1
                Case "L": myX = myX - 1
                Case Else

                    Err.Raise vbObjectError + 452, "Day03:GetWirePath", Fmt("Unknown direction '{0}'", myDirection)
                    
            End Select

            Dim myTotalSteps As Long
            myTotalSteps = myTotalSteps + 1

            Dim myCoords  As Point
            Set myCoords = Point.Make(myX, myY)

            If myWirePath.LacksKey(myCoords) Then

                myWirePath.AddByKey myCoords, New KvpTotalSteps
                
            End If
            
            myWirePath.Item(myCoords).AddByIndex myTotalSteps

        Next

    Next

    Set GetWirePath = myWirePath
    
End Function


Public Function GetIntersectionsOfWires _
( _
    ByRef ipWire1Path As KvpWireAsPath, _
    ByRef ipWire2Path As KvpWireAsPath _
) As KvpWirePathIntersections

    Dim myKey As Variant
    Dim myIntersections As KvpWirePathIntersections: Set myIntersections = New KvpWirePathIntersections
    
    For Each myKey In ipWire1Path
        
        DoEvents
        If ipWire2Path.HoldsKey(myKey) Then
        
            If myIntersections.LacksKey(myKey) Then
            
                myIntersections.AddByKey myKey, New KvpStepsW1W2
                
            End If
        
            'myIntersections.Item(myKey).AddByIndex GetCoordIntersectionList(ipWire1Path.Item(myKey), ipWire2Path.Item(myKey))
            
        End If

    Next
    
    Set GetIntersectionsOfWires = myIntersections
    
End Function


Public Function GetCoordIntersectionList(ByRef ipStepsW1 As KvpTotalSteps, ByRef ipStepsW2 As KvpTotalSteps) As KvpStepsW1W2
    
    Dim myStepsW1 As Variant
    Dim myIntersections As KvpStepsW1W2: Set myIntersections = New KvpStepsW1W2
    For Each myStepsW1 In ipStepsW1
        DoEvents
        Dim myStepsW2 As Variant
        For Each myStepsW2 In ipStepsW2
        
            DoEvents
            myIntersections.AddByIndex StepsW1W2.Make(ipStepsW1.Item(myStepsW1), ipStepsW2.Item(myStepsW2))
            
        Next
        
    Next
        
    Set GetCoordIntersectionList = myIntersections
    
End Function

Public Function GetManhattenDistance(ByVal ipCoord As String) As Long

    GetManhattenDistance = Abs(PointX(ipCoord)) + Abs(PointY(ipCoord))
   
End Function


Public Function GetMinManhattenDistance(ByRef ipWireIntersections As KvpWirePathIntersections) As Long

    Const MAX_LONG As Long = &H7FFFFFFF
    
    Dim myKey As Variant
    Dim myMin As Long: myMin = MAX_LONG
    Dim myManhattenDistance As Long
    For Each myKey In ipWireIntersections
        
        DoEvents
        Dim myCoord As String
        myCoord = myKey
        myManhattenDistance = GetManhattenDistance(myCoord)
        myMin = IIf(myMin < myManhattenDistance, myMin, myManhattenDistance)
        
    Next
    
    GetMinManhattenDistance = myMin
    
End Function


Public Function GetIntersectionWithMinimumSteps(ByRef ipWireIntersections As KvpWirePathIntersections)
    'KvpWirePathIntersections = key:String vs kvp of key:long by Value:string,IntersectionSteps
    Const MAX_LONG As Long = &H7FFFFFFF
    
    Dim myMin As Long: myMin = MAX_LONG
    
    Dim myWireIntersection As Variant
    For Each myWireIntersection In ipWireIntersections
    
        DoEvents
        Dim myStepsXStepsY As Variant
        Dim myKvp As Variant: Set myKvp = ipWireIntersections.Item(myWireIntersection)
        For Each myStepsXStepsY In myKvp.Kvp.GetKeys
            
            DoEvents
            Dim mySumSteps As Long
            mySumSteps = IntersectionStepsX(myKvp.Kvp.Item(myStepsXStepsY)) + IntersectionStepsY(myKvp.Kvp.Item(myStepsXStepsY))
            
            myMin = IIf(myMin < mySumSteps, myMin, mySumSteps)
        Next
    Next
    
    GetIntersectionWithMinimumSteps = myMin
    
End Function





