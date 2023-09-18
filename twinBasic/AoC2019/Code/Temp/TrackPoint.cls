VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrackPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
''@Exposed
'Option Explicit
'
'Public Enum TurnInstruction
'
'    Left = 0&
'    Right = 1&
'
'End Enum
'
'
'Public Enum Direction
'
'    Unknown = -1&
'    North = 0&
'    East = 1&
'    South = 2&
'    west = 3&
'
'End Enum
'
'
'Private Type Properties
'
'    ArriveColour                As PanelColour
'    DepartColour                As PanelColour
'    Step                        As Long
'    ArriveDirection             As Direction
'    Turn                        As TurnInstruction
'    DepartDirection             As Direction
'    CoordX                      As Long
'    CoordY                      As Long
'
'End Type
'
'Private p                       As Properties
'
'
'Public Property Let ArriveColour(ByVal ipValue As PanelColour)
'    p.ArriveColour = ipValue
'End Property
'
'Public Property Get ArriveColour() As PanelColour
'    ArriveColour = p.ArriveColour
'End Property
'
'
'Public Property Let DepartColour(ByVal ipValue As PanelColour)
'    p.DepartColour = ipValue
'End Property
'
'Public Property Get DepartColour() As PanelColour
'    DepartColour = p.DepartColour
'End Property
'
'
'Public Property Let Step(ByVal ipValue As Long)
'    p.Step = ipValue
'End Property
'
'Public Property Get Step() As Long
'    Step = p.Step
'End Property
'
'
'Public Property Let ArriveDirection(ByVal ipValue As Direction)
'    p.ArriveDirection = ipValue
'End Property
'
'Public Property Get ArriveDirection() As Direction
'    ArriveDirection = p.ArriveDirection
'End Property
'
'
'Public Property Get Turn() As TurnInstruction
'    Turn = p.Turn
'End Property
'
'Public Property Let Turn(ByVal ipValue As TurnInstruction)
'    p.Turn = ipValue
'End Property
'
'
'Public Property Get DepartDirection() As Direction
'    DepartDirection = p.DepartDirection
'End Property
'
'Public Property Let DepartDirection(ByVal ipValue As Direction)
'    p.DepartDirection = ipValue
'End Property
'
'
'Public Property Get CoordX() As Long
'    CoordX = p.CoordX
'End Property
'
'Public Property Let CoordX(ByVal ipValue As Long)
'    p.CoordX = ipValue
'End Property
'
'
'Public Property Get CoordY() As Long
'    CoordY = p.CoordY
'End Property
'
'Public Property Let CoordY(ByVal ipValue As Long)
'    p.CoordY = ipValue
'End Property
'
''@Ignore ProcedureNotUsed
'Public Function Clone() As TrackPoint
'
'    Dim myNewTp As TrackPoint: Set myNewTp = New TrackPoint
'
'    With myNewTp
'
'        .ArriveColour = Me.ArriveColour
'        .ArriveDirection = Me.ArriveDirection
'        .CoordX = Me.CoordX
'        .CoordY = Me.CoordY
'        .DepartColour = Me.DepartColour
'        .DepartDirection = Me.DepartDirection
'        .Step = Me.Step + 1
'        .Turn = Me.Turn
'
'    End With
'
'    Set Clone = myNewTp
'
'End Function
'
'
'Public Function Coords() As String
'    Coords = "(" & CStr(Me.CoordX) & "," & CStr(Me.CoordY) & ")"
'End Function
'
'
'
'
