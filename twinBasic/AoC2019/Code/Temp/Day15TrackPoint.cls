VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Day15TrackPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Exposed
'@PredeclaredId
''@Exposed
''@PredeclaredId
'Option Explicit
'
'Private Type properties
'
'    Steps                       As Long
'    IsType                      As TileType
'    X                           As Long
'    Y                           As Long
'    Explored                    As Boolean
'
'End Type
'
'Private p                       As properties
'
'
'Public Function Debutante(ByVal ipX As Long, ByVal ipY As Long, ByVal ipType As Long, ByVal ipSteps As Long) As Day15TrackPoint
'
'    With New Day15TrackPoint
'
'        Set Debutante = .Self(ipX, ipY, ipType, ipSteps)
'
'    End With
'
'End Function
'
'
'Public Function Self(ByVal ipX As Long, ByVal ipY As Long, ByVal ipType As Long, ByVal ipSteps As Long) As Day15TrackPoint
'
'    p.X = ipX
'    p.Y = ipY
'    p.IsType = ipType
'    p.Steps = ipSteps
'    p.Explored = False
'    Set Self = Me
'
'End Function
'
'
'Public Property Get X() As Long
'    X = p.X
'End Property
'
'
'Public Property Get Y() As Long
'    Y = p.Y
'End Property
'
'
'Public Property Get IsType() As TileType
'    IsType = p.IsType
'End Property
'
'
'Public Property Let IsType(ByVal ipValue As TileType)
'    p.IsType = ipValue
'End Property
'
'
'Public Property Get Steps() As Long
'    Steps = p.Steps
'End Property
'
'Public Property Let Steps(ByVal ipValue As Long)
'    p.Steps = ipValue
'End Property
'
'
'Public Property Get Explored() As Long
'    Explored = p.Explored
'End Property
'
'Public Property Let Explored(ByVal ipValue As Long)
'    p.Explored = ipValue
'End Property
'
'
'Public Function NeighbourCoords(ByVal ipWay As Way) As String
'
'    Dim myX As Long
'    myX = p.X
'    Dim myY As Long
'    myY = p.Y
'    Select Case ipWay
'
'        Case Way.North: myY = myY + 1
'        Case Way.South: myY = myY - 1
'        Case Way.West: myX = myX + 1
'        Case Way.East: myX = myX - 1
'
'    End Select
'
'    NeighbourCoords = Coords(myX, myY)
'
'End Function
'
'
'Public Function Coords(Optional ByVal ipX As Variant, Optional ByVal ipY As Variant) As String
'
'    Dim myX As Long
'    Dim myY As Long
'    If IsMissing(ipX) Then
'
'        myX = p.X
'        myY = p.Y
'
'    Else
'
'        myX = ipX
'        myY = ipY
'
'    End If
'
'    Coords = Layout.Format("({0},{1})", Array(myX, myY))
'
'End Function
'
