VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PaintingRobot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Exposed
'@PredeclaredId
''@PredeclaredId
'Option Explicit
''@Exposed
'
'Private Enum OutputIndex
'
'    PaintColour = 0&
'    TurnInstruction = 1&
'
'End Enum
'
'
'Private Type Properties
'
'    Track                   As Kvp
'
'End Type
'
'Private p                   As Properties
'
'
'Private Type State
'
'    Computer                As IntComputer
'    Board                   As ExcelBoard
'
'End Type
'
'Private s                    As State
'
'
'
'Public Function Debutante _
'( _
'    Optional ByVal ipStartX As Long = 0, _
'    Optional ByVal ipStartY As Long = 0, _
'    Optional ByVal ipDirection As Direction = Direction.North _
') As PaintingRobot
'
'    With New PaintingRobot
'
'        Set Debutante = .Self(ipStartX, ipStartY, ipDirection)
'
'    End With
'
'End Function
'
'
'Public Function Self _
'( _
'    ByVal ipStartX As Long, _
'    ByVal ipStartY As Long, _
'    ByVal ipDirection As Direction _
') As PaintingRobot
'
'    Set s.Computer = New IntComputer
'    Set p.Track = New KvpOD
'    Set s.Board = New ExcelBoard
'
'    Dim myTp As TrackPoint: Set myTp = New TrackPoint
'
'    myTp.ArriveColour = Black
'    myTp.DepartDirection = ipDirection
'    myTp.ArriveDirection = Unknown
'    myTp.CoordX = ipStartX
'    myTp.CoordY = ipStartY
'    myTp.Step = 0
'    UpdateTrack myTp
'
'    Set Self = Me
'
'End Function
'
'
'Public Property Set Program(ByVal ipProgram As KvpOD)
'
'    Set s.Computer.Program = ipProgram
'
'End Property
'
'
'Public Property Get Program() As KvpOD
'
'    Set Program = s.Computer.Program
'
'End Property
'
'
'Public Sub Run(Optional ByVal ipInput As KvpOD)
'
'    s.Computer.OutputMode = HaltOnOutput
'    Dim myBoard As ExcelBoard
'    Set myBoard = New ExcelBoard
'
'    Dim myInput As Kvp
'    Set myInput = ipInput.Clone
'    Dim myCurrentTp As TrackPoint
'    'Debug.Print p.Track.GetKeysAsString
'    'Debug.Print p.Track.GetFirst.Value.GetKeysAsString
'
'    Set myCurrentTp = p.Track.GetFirst.Value.GetFirst.Value
'    myBoard.PlaceTile myCurrentTp.CoordX, myCurrentTp.CoordY, ipDirection:=myCurrentTp.DepartDirection
'    'Set p.Track = New KvpOD
'    'Debug.Print "Track count is ", p.Track.Count
'    Do
'        ' The computer run is in two parts.
'        ' The first part requires an input, which is the colour
'        ' of the current position; This colour is the arrivecolour of the
'        ' last trackpoint in p.Track which is captured after the current trackpoint is updated.
'
'        s.Computer.Run myInput
'        If s.Computer.RunHasCompleted Then Exit Do
'        'Debug.Print s.Computer.GetOutput.GetValuesAsString
'
'        myCurrentTp.DepartColour = s.Computer.GetOutput.GetFirst.Value 'GetItem(OutputIndex.PaintColour)
'
'        ' No input is required for part 2 of the run
'        s.Computer.Run ' No input required for second run
'        If s.Computer.RunHasCompleted Then Exit Do
'        'Debug.Print s.Computer.GetOutput.GetValuesAsString
'        myCurrentTp.Turn = s.Computer.GetOutput.GetFirst.Value  'Item(OutputIndex.TurnInstruction)
'
'        Set myCurrentTp = TurnThenMove(myCurrentTp)
'        Set myInput = MakeKvp(myCurrentTp.ArriveColour)
'
'        '@Ignore FunctionReturnValueDiscarded
'        DoEvents
'
'    Loop
'    Debug.Print p.Track.GetKeysAsString
'
'End Sub
'
'Private Function TurnThenMove(ByRef ipLastTp As TrackPoint) As TrackPoint
'
'    Dim myNewTp As TrackPoint
'    Set myNewTp = ipLastTp.Clone
'
'    myNewTp.ArriveDirection = ipLastTp.DepartDirection
'    'iplast is now complete so we can store it on the track
'    UpdateTrack ipLastTp
'
'
'    With myNewTp
'
'        'now turn and move
'        Select Case .ArriveDirection
'
'            Case North:
'
'                If .Turn = Left Then
'
'                    .DepartDirection = west
'                    .CoordX = .CoordX - 1
'
'                Else
'
'                    .DepartDirection = East
'                    .CoordX = .CoordX + 1
'
'                End If
'
'
'            Case South
'
'                If .Turn = Left Then
'
'                    .DepartDirection = East
'                    .CoordX = .CoordX + 1
'
'                Else
'
'                    .DepartDirection = west
'                    .CoordX = .CoordX - 1
'
'                End If
'
'
'            Case East
'
'                If .Turn = Left Then
'
'                    .DepartDirection = North
'                    .CoordY = .CoordY + 1
'
'                Else
'
'                    .DepartDirection = South
'                    .CoordY = .CoordY - 1
'
'                End If
'
'            Case west
'
'                If .Turn = Left Then
'
'                    .DepartDirection = South
'                    .CoordY = .CoordY - 1
'
'                Else
'
'                    .DepartDirection = North
'                    .CoordY = .CoordY + 1
'
'                End If
'
'            Case Else
'
'                Debug.Print "UpdateDirectionAndCoordinates: Illegal direction " & .Turn
'
'        End Select
'
'        .ArriveColour = ColourFromRobotEye(myNewTp)
'        ShowPositionInExcel ipLastTp.CoordX, ipLastTp.CoordY, ipDirection:=.DepartDirection
'        ShowPositionInExcel .CoordX, .CoordY, ipDirection:=.ArriveDirection
'        ShowPositionInExcel ipLastTp.CoordX, ipLastTp.CoordY, ipColour:=.DepartColour
'
'    End With
'
'
'    Set TurnThenMove = myNewTp
'
'End Function
'
'Public Function ColourFromRobotEye(ByVal ipTp As TrackPoint) As PanelColour
'
'    ColourFromRobotEye = PanelColour.Black
'    If p.Track Is Nothing Then Exit Function
'    If p.Track.Count = 0 Then Exit Function
'    If p.Track.LacksKey(ipTp.Coords) Then Exit Function
'
'    ColourFromRobotEye = p.Track.Item(ipTp.Coords).GetLast.Value.DepartColour
'
'End Function
'
'
''@Ignore FunctionReturnValueNotUsed
'Public Function Track() As KvpOD
'    Set Track = p.Track
'End Function
'
'
'Public Sub UpdateTrack(ByVal ipTp As TrackPoint)
'
'    If p.Track.LacksKey(ipTp.Coords) Then
'
'        p.Track.AddByKey ipTp.Coords, New KvpOD
'
'    End If
'
'    p.Track.Item(ipTp.Coords).AddByIndex ipTp
'
'End Sub
'
'
'
'
'
'
'
'
'
'
'
'


