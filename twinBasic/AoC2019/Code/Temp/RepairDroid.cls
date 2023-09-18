VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepairDroid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
''@Exposed
''@PredeclaredId
'Option Explicit
'
''@Ignore EncapsulatePublicField
''Public WithEvents myKeyboard As InputCollector
'
'
'Public Enum Way
'
'    North = 1
'    South = 2
'    West = 3
'    East = 4
'
'End Enum
'
'
'Public Enum TileType
'
'    Wall = 0
'    Movable = 1
'    Oxygen = 2
'
'End Enum
'
'
'Private Type properties
'
'    Computer                        As IntComputer
'
'End Type
'
'Private p                           As properties
'
'
'Private Type State
'
'    Tp                              As Day15TrackPoint
'    Board                           As ExcelBoard
'    Known                           As Kvp
'    MoveResult                      As Kvp
'    Exploring                       As Kvp
'
'End Type
'
'Private s                           As State
'
'
'Public Function Debutante(ByVal ipX As Long, ByVal ipY As Long, ByVal ipType As TileType, ByVal ipSteps As Long) As RepairDroid
'
'    With New RepairDroid
'
'        Set Debutante = .Self(ipX, ipY, ipType, ipSteps)
'
'    End With
'
'End Function
'
'
'Public Function Self(ByVal ipX As Long, ByVal ipY As Long, ByVal ipType As TileType, ByVal ipSteps As Long) As RepairDroid
'
'    Set p.Computer = New IntComputer
'    p.Computer.OutputMode = HaltOnOutput
'
'    Set s.Known = New KvpOD
'    Set s.Exploring = New KvpOD
'
'    Set s.Tp = Day15TrackPoint.Debutante(ipX, ipY, ipType, ipSteps)
'    s.Tp.Explored = False
'    'AddToMap s.Tp
'
'    Set s.Board = New ExcelBoard
'
'    s.Board.PlaceTile s.Tp.X, s.Tp.Y, "D"
'
''    Set s.MoveResult = New KvpOD
''    s.MoveResult.AddByKeyFromArrays _
''        Array(TileType.Wall, TileType.moveable, TileType.Oxygen), _
''        Split("#,.,O", ",")
'
'    Set Self = Me
'
'End Function
'
'Public Sub Run()
'
'    Dim myInput As Kvp
'    Set myInput = MakeKvp(Way.North)
'
'    Dim myExploring                     As Boolean
'    myExploring = True
'
'
'
'    Do While myExploring
'
'        Dim myDroid As Day15TrackPoint
'        Set myDroid = Day15TrackPoint.Debutante(s.Tp.X, s.Tp.Y, s.Tp.IsType, s.Tp.Steps)
'    'MoveDroid
'        p.Computer.Run myInput
'        If p.Computer.RunHasCompleted Then Exit Do
'        s.Tp.IsType = p.Computer.GetOutput.GetLast.Value
'
'
'        Select Case myOutput
'
'            Case TileType.Wall
'
'                s.Board.PlaceTile s.Tp.X, s.Tp.Y, "W"
'
'            Case TileType.Oxygen
'
'            Case TileType.Movable
'
'
'     If Not s.Tp.Explored Then
'
'        ' ExploreUnknownNeighbours
'
'     Else
'
'    ' If AnyUnknownNeighboursRemain Then
'
''Loop
'End If
'
'End Sub
'
'Public Sub AddToMap(ByRef ipTp As Day15TrackPoint, ByVal ipMap As KvpOD)
'
'    ' Kvp for each step
'    If ipMap.LacksKey(ipTp.Steps) Then
'
'        ipMap.AddByKey ipTp.Steps, New KvpOD
'
'    End If
'
'
'    'Kvp for each X coordinate
'    Dim myXKvp As Kvp
'    Set myXKvp = ipMap.Item(ipTp.Steps)
'    If myXKvp.LacksKey(ipTp.X) Then
'
'        myXKvp.AddByKey ipTp.X, New KvpOD
'
'    End If
'    ' Kvp for each Y coordinate
'    'At the Y level we can only hold trackpoints
'    Dim myYKvp As Kvp
'    Set myYKvp = myXKvp.Item(ipTp.X)
'    myYKvp.AddByKey ipTp.Y, ipTp
'
'End Sub
'
'
'
'
'Public Function TilesWithUnknownNeighboursExist() As Boolean
'
'    Dim myReturn As Boolean
'    myReturn = False
'
'    Dim myItem As Variant
'    For Each myItem In s.Map
'
'        Dim myTrackpoint As Day15TrackPoint
'        Set myTrackpoint = myItem
'        If Not myTrackpoint.Explored Then
'
'            myReturn = True
'            Exit For
'
'        End If
'
'    Next
'
'    TilesWithUnknownNeighboursExist = myReturn
'
'End Function
'Public Sub TestNewTile(ByVal ipWay As Way)
'
'
'
'    p.Computer.Run MakeKvp(ipWay)
'    If p.Computer.RunHasCompleted Then Exit Sub
'
'    Dim myOutput As Long
'    myOutput = p.Computer.GetOutput.GetLast.Value
'
'   ' we now need to determine
'
'    If myOutput = TileType.Wall Then
'
'        s.Tp.IsType = Wall
'        s.Board.PlaceTile s.Tp.X, s.Tp.Y, "#"
'
'    End If
'
'    Dim myTp As Day15TrackPoint
'    Set myTp = New Day15TrackPoint
'
'
'
'
'
'    If myOutput = TileType.Oxygen Then
'
'
''        Select Case Way
''
''            Case north:
''            Case South
''            Ca
'
'    Else
'
'      '  s.Tp.TrackPointTileType=
'
'      End If
'
'
'
'End Sub
'
'
'Public Property Set Program(ByVal ipValue As KvpOD)
'    Set p.Computer.Program = ipValue
'End Property
'
'
''Public Sub UpdateBoard(ByVal ipTp As Day15TrackPoint)
''
''End Sub
'
''    Dim myKeyToWay As KvpOD: Set myKeyToWay = New KvpOD
''    myKeyToWay.AddByKeyFromArrays _
''        Array(vbKeyUp, vbKeyDown, vbKeyRight, vbKeyLeft), _
''        Array(Way.North, Way.South, Way.East, Way.West)
''
''    Dim myKeyboard As InputCollector
''    Set myKeyboard = New InputCollector
''    Set myKeyboard.KeyToWayMap = myKeyToWay
'
''Private Sub myKeyboard_MoveToKnownTile(ByVal ipCoords As String)
''    MoveToKnownTile ipCoords
''End Sub
''
''Private Sub myKeyboard_TestNewTile(ByVal ipWay As Way)
''    TestNewTile ipWay
''End Sub
'


