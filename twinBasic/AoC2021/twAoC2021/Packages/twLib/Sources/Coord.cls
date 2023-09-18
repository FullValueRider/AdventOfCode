VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Const MT_NAME                               As string = "Coord"
Private Const FOUR_POINTS_TURN_ANGLE        As Long = 90
Private Const EIGHT_POINTS_TURN_ANGLE       As Long = 45
Private Const EIGHT_COMPASS_POINTS          As Long = 8
Private Const FOUR_COMPASS_POINTS           As Long = 4

Private Type XYMovement

    YAxis                                   As long
    XAxis                                   As long

End Type


Public Type Bounds

    MinX                                    As Long
    MinY                                    As Long
    MaxX                                    As Long
    MaxY                                    As Long
    
End Type

Public Enum Direction

    North = 0
    East = 90
    South = 180
    West = 270
    
    NW = 315
    NE = 45
    SE = 135
    SW = 225
    
    Forward = 1
    Backward = 2
    Left = 3
    Right = 4
    
End Enum

' Public Enum Turn

'     Right
'     Left
    
' End Enum


Private Type Properties

    CurrentX                                As Long
    CurrentY                                As Long
    Direction                               As Direction
    TurnAngle                               As long ' 45 or 90 degrees
    RecordFirstRepeatVisit                  As boolean
    SiteOfFirstRepeatVisit                  As string
    
End Type

Private p As Properties


Private Type State

	OriginX                                 As Long
    OriginY                                 As Long
    
    DirectionMap                            As lookup
    RightTurn                               As Lookup
    LeftTurn                                As Lookup
    Movement                                As XYMovement
    MovementMap                             As lookup
    Visited                                 As Hkvp
    Blocked                                 As Hkvp ' location to which we cannot move
    Track                                   As Seq
    IsBounded                               As Boolean
    Limits                                  As Bounds
    
End Type

Private s                               As state

'@DefaultMember
Public Function Deb() As Coord
	With New Coord
        Return .ConstructInstance(0, 0, FOUR_COMPASS_POINTS)
    End With
End Function

Public Function Deb(ByVal ipX As Long, ByVal ipY As Long) As Coord
    With New Coord
        Return .ConstructInstance(CLng(ipX), CLng(ipY), FOUR_COMPASS_POINTS)
    End With
End Function

Public Function Deb(ByVal ipX As Long, ByVal ipY As Long, ByVal ipCompassPoints As Long) As Coord
    With New Coord
        Return .ConstructInstance(CLng(ipX), CLng(ipY), ipCompassPoints)
    End With
End Function

Public Function Deb(ByVal ipXYCoord As String) As Coord
    
    Dim myCoord As Variant = Split(ipXYCoord, Char.twComma)
    Return Deb(CLng(myCoord(0)), CLng(myCoord(1)), FOUR_COMPASS_POINTS)

End Function

Public Function Deb(ByVal ipXYCoord As String, ByVal ipCompassPoints As Long) As Coord
    
    Dim myCoord As Variant = Split(ipXYCoord, Char.twComma)
    Return Deb(CLng(myCoord(0)), CLng(myCoord(1)), ipCompassPoints)

End Function


Public Function Deb(ByVal ipKVPair As KVPair) As Coord

    Return Deb(CLng(ipKVPair.Key(0)), CLng(ipKVPair.Item(0)), FOUR_COMPASS_POINTS)
    
End Function

Public Function Deb(ByVal ipKVPair As KVPair, ByVal ipCompassPoints As Long) As Coord

    Return Deb(CLng(ipKVPair.Key(0)), CLng(ipKVPair.Item(0)), ipCompassPoints)
    
End Function


Public Function ConstructInstance(ByVal ipX As Long, ByVal ipY As Long, ByVal ipCompassPoints As Long) As Coord

    p.CurrentX = ipX
    p.CurrentY = ipY
    
    s.OriginX = ipX
    s.OriginY = ipY
    
    p.Direction = Direction.North
    s.Movement.XAxis = 0
    s.Movement.YAxis = 1
    
    If ipCompassPoints <> FOUR_COMPASS_POINTS And ipCompassPoints <> EIGHT_COMPASS_POINTS Then
        Err.Raise 17 + vbObjectError, "twLib.Coord.ConstructInstance", "Compass points must be 4 or 8"
    End If
    
    p.TurnAngle = IIf(ipCompassPoints = 4, FOUR_POINTS_TURN_ANGLE, EIGHT_POINTS_TURN_ANGLE)
    
    Set s.DirectionMap = Lookup.Deb
    
    With s.DirectionMap
    
        .Add "n", Direction.North
        .Add "north", Direction.North
        .Add "u", Direction.North
        .Add "up", Direction.North
        .Add "^", Direction.North
        
        .Add "s", Direction.South
        .Add "south", Direction.South
        .Add "d", Direction.South
        .Add "down", Direction.South
        .Add "v", Direction.South
        
        
        .Add "w", Direction.West
        .Add "west", Direction.West
        .Add "<", Direction.West
        
        .Add "e", Direction.East
        .Add "east", Direction.East
        .Add ">", Direction.East
        
        
        .Add "nw", Direction.NW
        .Add "ne", Direction.NE
        .Add "se", Direction.SE
        .Add "sw", Direction.SW
        
        .Add "f", Direction.Forward
        .Add "forward", Direction.Forward
        .Add "forwards", Direction.Forward
       
        .Add "b", Direction.Backward
        .Add "back", Direction.Backward
        .Add "backward", Direction.Backward
        .Add "backwards", Direction.Backward
        
        .Add "l", Direction.Left
        .Add "left", Direction.Left
        
        .Add "r", Direction.Right
        .Add "right", Direction.Right
       
    End With
    
    Set s.LeftTurn = Lookup.Deb
    With s.LeftTurn
    
        .Add "0,1", Array(-1, 0)
        .Add "-1,0", Array(0, -1)
        .Add "0,-1", Array(1, 0)
        .Add "1,0", Array(0, 1)
        
    End With
    
    Set s.RightTurn = Lookup.Deb
    With s.RightTurn
    	
        .Add "0,1", Array(1, 0)
        .Add "1,0", Array(0, -1)
        .Add "0,-1", Array(-1, 0)
        .Add "-1,0", Array(0, 1)
        
    End With
    
    Set s.MovementMap = Lookup.Deb
    With s.MovementMap
    
        .Add Direction.North, Array(0, 1)
        .Add Direction.NE, Array(1, 1)
        .Add Direction.East, Array(1, 0)
        .Add Direction.SE, Array(1, -1)
        .Add Direction.South, Array(0, -1)
        .Add Direction.SW, Array(-1, -1)
        .Add Direction.West, Array(-1, 0)
        .Add Direction.NW, Array(-1, 1)
        
    End With
    
    p.RecordFirstRepeatVisit = False
    p.SiteOfFirstRepeatVisit = vbNullString
    Set s.Track = Seq.Deb
    Set s.Visited = Hkvp.Deb
    s.IsBounded = False
    
    Return Me
    
End Function

Public Property Get RecordFirstRepeatVisit() As Boolean
    Return p.RecordFirstRepeatVisit
End Property

Public Property Let RecordFirstRepeatVisit(ByVal ipBoolean As Boolean)
	p.RecordFirstRepeatVisit = ipBoolean
End Property

' Public Function GetDirection(ByVal ipDirection As String) As Direction
'     Return s.DirectionMap.Item(ipDirection)
' End Function

Public Function SetBlocked(ByVal ipBlocked As Hkvp) As Coord
    Set s.Blocked = ipBlocked
    Return Me
End Function

Public Function IsAtOrigin() As Boolean
    Return (s.OriginX = p.CurrentX) And (s.OriginY = p.CurrentY)
End Function

Public Function Reset() As Coord
    p.CurrentX = s.OriginX
    p.CurrentY = s.OriginY
    Return Me
End Function

Public Property Get X() As Long
    X = p.CurrentX
End Property

Public Property Let X(ByVal ipX As Long)
    p.CurrentX = ipX
End Property

Public Property Get Y() As Long
    Y = p.CurrentY
End Property

Public Property Let Y(ByVal ipY As Long)
    p.CurrentY = ipY
End Property


Public Property Get ToString() As String
    Return CStr(p.CurrentX) & Char.twComma & CStr(p.CurrentY)
End Property


Public Function Clone() As Coord
    Return Deb(s.OriginX, s.OriginY, CLng(IIf(p.TurnAngle = 90, FOUR_COMPASS_POINTS, EIGHT_COMPASS_POINTS))).MoveTo(p.CurrentX, p.CurrentY)
End Function


Public Function MoveTo(ByVal ipX As Long, ipY As Long) As Coord

    If s.Blocked IsNot Nothing Then
    
        If s.Blocked.HoldsKey(CStr(ipX) & Char.twComma & CStr(ipY)) Then
            Return Me
        End If
        
    End If
    
    p.CurrentX = ipX
    p.CurrentY = ipY
    If s.IsBounded Then
        EnforceBounds
    End If
        Return Me
End Function


Public Sub Move(ByVal ipMoveDistance As Long = 1)

    Dim mySteps As Long
    Dim myNewX As Long
    Dim myNewY As Long
    
    For mySteps = 1 To VBA.Abs(ipMoveDistance)
    
        If ipMoveDistance > 0 Then
    
            myNewX = p.CurrentX + s.Movement.XAxis
            myNewY = p.CurrentY + s.Movement.YAxis
            
            
        Else
        
            myNewX = p.CurrentX - s.Movement.XAxis
            myNewY = p.CurrentY = s.Movement.YAxis
        
        
        End If
        
        If s.Blocked Is Nothing Then
        
            p.CurrentX = myNewX
            p.CurrentY = myNewY
            
        Else
        
            If s.Blocked.LacksKey(CStr(myNewX) & Char.twComma & CStr(myNewY)) Then
                
                p.CurrentX = myNewX
                p.CurrentY = myNewY
                
            End If
            
            
        End If
        
        If s.IsBounded Then
            EnforceBounds
        End If
        
        Dim mylocation As String = ToString
        
        s.Track.Add mylocation
        If s.Visited.HoldsKey(mylocation) Then
            
            s.Visited.Item(mylocation) += 1
            
            If p.RecordFirstRepeatVisit Then
        
                p.SiteOfFirstRepeatVisit = mylocation
                p.RecordFirstRepeatVisit = False
            
            
            End If
            
            
        Else
        
            s.Visited.Add mylocation, 1
            
            
        End If
        
    Next

End Sub


Private Sub EnforceBounds()

    If p.CurrentX > s.Limits.MaxX Then p.CurrentX = s.Limits.MaxX
    If p.CurrentX < s.Limits.MinX Then p.CurrentX = s.Limits.MinX
    If p.CurrentY > s.Limits.MaxY Then p.CurrentY = s.Limits.MaxY
    If p.CurrentY < s.Limits.MinY Then p.CurrentY = s.Limits.MinY
    
End Sub


Public Function SetBounds(ByVal ipMinX As Long, ByVal ipMinY As Long, ByVal ipMaxX As Long, ByVal ipMaxY As Long) As Coord
	
        s.IsBounded = True
        
        s.Limits.MinX = ipMinX
        s.Limits.MinY = ipMinY
        s.Limits.MaxX = ipMaxX
        s.Limits.MaxY = ipMaxY
    
        Return Me
        
End Function


Public Property Get IsBounded() As Boolean
    Return s.IsBounded
End Property


Public Function Proceed(ByVal ipDirection As String, Optional ByVal ipdistance As Long = 1) As Coord

    ProceedByDirection s.DirectionMap.Item(VBA.LCase$(ipDirection)), ipdistance
    
    Return Me
    
End Function


Private Function ProceedByDirection(ByVal ipDirection As Direction, ByVal ipdistance As Long) As Coord
    
    Dim mymovement As Variant
    
    Select Case ipDirection
    
        Case North, NE, East, SE, South, SW, West, NW
        
            mymovement = s.MovementMap.Item(ipDirection)
            s.Movement.XAxis = mymovement(0)
            s.Movement.YAxis = mymovement(1)
    
        ' Fowrward, backward, left and right require movement relative 
        ' to the current forward direction
        
        Case Forward
        
            'no need to do anything as s.movement defines the forward direction
        
        Case Backward
        
            'reverse movement
            
            s.Movement.XAxis *= -1
            s.Movement.YAxis *= -1
            
        ' left and right need lookups based on current s.movement
        Case Direction.Left
        
            mymovement = s.LeftTurn.Item(CStr(s.Movement.XAxis) & Char.twComma & CStr(s.Movement.YAxis))
            s.Movement.XAxis = mymovement(0)
            s.Movement.YAxis = mymovement(1)
            
        
        Case Direction.Right
        
            mymovement = s.RightTurn.Item(CStr(s.Movement.XAxis) & Char.twComma & CStr(s.Movement.YAxis))
            s.Movement.XAxis = mymovement(0)
            s.Movement.YAxis = mymovement(1)
        
        
    End Select
    
    Move ipdistance
    
    Return Me
    
End Function


Public Function Manhatten() As Long
    Return Abs(p.CurrentX - s.OriginX) + Abs(p.CurrentY - s.OriginY)
End Function


Public Function Manhatten(ByVal ipLocation As String) As Long
    Dim myLocation As Variant = VBA.Split(ipLocation, Char.twComma)
    Return Manhatten(CLng(myLocation(0)), CLng(myLocation(1)))
End Function

Public Function Manhatten(ByVal ipX As Long, ipY As Long) As Long
    Return Abs(ipX - s.OriginX) + Abs(ipY - s.OriginY)
End Function

Public Function MapHeading() As Double
    ' Heading when the YAxis is positive moving up
    ' Debug.Print "Need to fix Heading"
    ' Dim myDx As Double
    ' myDx = CDbl(p.CurrentX - s.OriginX)
    
    ' Dim myDy As Double
    ' myDy = CDbl(p.CurrentY - s.OriginY)
    
    ' Dim myAngle As Double
    ' myAngle = Maths.ATan2(myDy, myDx) * Maths.RadiansToDegrees
    
    ' If myAngle < 0 Then
    
    '   '  myAngle = myAngle + p.HeadingDatum + 360   ' This is simular to doing
    '                                 ' 360 Math.Atan2(y1 - y2, x1 - x2) * (180 / Math.PI)
                                    
    ' End If
    
    ' Heading = myAngle
   
End Function

Public Function VDUHeading() As Double
    ' Heading when the Y axis is positive moving down
End Function

Public Property Get TurnAngle() As Long
    Return p.TurnAngle
End Property

Public Property Get CompassPoints() As Long
    Return IIf(p.TurnAngle = 90, 4, 8)
End Property


Public Function SiteOfFirstRepeatVisit() As String
    Return p.SiteOfFirstRepeatVisit
End Function
' Public Sub Turn(ByVal ipTurn As String)


'     Select Case ipTurn
    
'         Case "R", "r", ">"
        
'             p.Direction += p.TurnAngle
                
'             If p.Direction >= 360 Then
            
'                 p.Direction -= 360
                
'             End If
            
            
        
'         Case "L", "l", "<"
        
'             p.Direction -= p.TurnAngle
            
'             If p.Direction < 0 Then
                
'                 p.Direction += 360
'             End If
            
            
'         Case Else
            
'                 Err.Raise 17 + vbObjectError, "twLib.Coord.Turn", "The turn parameter should be one of 'R,r,L,l'"
        
'     End Select
    
'     Select Case p.Direction
    
'         Case 0
        
'             s.Movement.XAxis = 0
'             s.Movement.YAxis = 1
            
            
'         Case 45
        
'             s.Movement.XAxis = 1
'             s.Movement.YAxis = 1
            
            
'         Case 90
        
'             s.Movement.XAxis = 1
'             s.Movement.YAxis = 0
            
            
'         Case 135
        
'             s.Movement.XAxis = 1
'             s.Movement.YAxis = -1
            
            
'         Case 180
        
'             s.Movement.XAxis = 0
'             s.Movement.YAxis = -1
            
            
'         Case 225
        
'             s.Movement.XAxis = -1
'             s.Movement.YAxis = -1
            
            
'         Case 270
        
'             s.Movement.XAxis = -1
'             s.Movement.YAxis = 0
            
            
'         Case 315
        
'             s.Movement.XAxis = -1
'             s.Movement.YAxis = 1
            
        
'     End Select

' End Sub