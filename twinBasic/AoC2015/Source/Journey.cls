VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Journey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Private Const NORTH                 As String = "^"
Private Const SOUTH                 As String = "v"
Private Const WEST                  As String = "<"
Private Const EAST                  As String = ">"

'Public Enum Way
'
'    North = 0
'    South = 1
'    East = 2
'    West = 3
'
'End Enum

Public Enum Move

    GoNorth = 1
    GoSouth = -1
    GoEast = 1
    GoWest = -1

End Enum

Private Type Properties

'    ArriveValue                     As String
'    DepartValue                     As String
    x                               As Long
    y                               As Long
'    ArriveWay                       As Long
'    DepartWay                       As Long
'    Steps                           As Long
'    PrevTp                          As Trackpoint
'    NextTp                          As Trackpoint
    Path                            As Kvp 'record of the sequence of x,y coords visited
    Visits                          As Kvp ' record of how many times each x,y coord was visited

End Type

Private p                           As Properties


Public Function Make(ByVal ipInstructions As Kvp) As Journey

    With New Journey

        Set Make = .Self(ipInstructions)

    End With

End Function


Public Function Self(ByVal ipInstructions As Kvp) As Journey

    p.x = 0
    p.y = 0
    
    Dim myPair As KVPair
    For Each myPair In ipInstructions
    
        Dim myCoords As String
        myCoords = NextHouse(myPair.Value)
        UpdatePath myCoords
        UpdateVisits myCoords
        
    Next

    Set Self = Me

End Function

Public Function NextHouse(ByVal ipInstruction As String) As String
    
    Select Case ipInstruction
    
        Case NORTH: p.y = p.y + Move.GoNorth
        
        Case SOUTH: p.y = p.y + Move.GoSouth
        
        Case EAST: p.x = p.x + Move.GoEast
        
        Case WEST: p.x = p.x + Move.GoWest

    End Select

    NextHouse = CStr(p.x) & "," & CStr(p.y)

End Function


Public Sub UpdatePath(ByVal ipCoords As String)

    ' Path is a Kvp of Key:Count of instructions vs Value: Coordinates as string
    
    If p.Path Is Nothing Then
    
        Set p.Path = New Kvp
        p.Path.SetFirstIndexAsLong 1&
        p.Path.AddByIndex "0,0"
    
    End If

    p.Path.AddByIndex ipCoords
           
End Sub

Public Sub UpdateVisits(ByVal ipCoords As String)

    ' Track is a Kvp of Key:Coordinates as string vs Value: Kvp of Instruction
    ' count when coordinates were visited
    ' i.e. When a house is visited more than once we know at what instruction
    ' steps the visits were
    
    Static myInstructionCount As Long

    If p.Visits Is Nothing Then
    
        myInstructionCount = 0
        Set p.Visits = New Kvp
        p.Visits.AddByKey "0,0", New Kvp
        p.Visits.GetItem("0,0").AddByIndex myInstructionCount
        
        
        
    End If
    
    myInstructionCount = myInstructionCount + 1
    
    If p.Visits.LacksKey(ipCoords) Then
    
        p.Visits.AddByKey ipCoords, New Kvp
        
    End If
    
    p.Visits.GetItem(ipCoords).AddByIndex myInstructionCount
    
End Sub

'Public Property Get x() As Long
'    x = p.x
'End Property
'
'
'Public Property Get y() As Long
'    y = p.y
'End Property
'
'
'Public Function GetPath() As Kvp
'    Set GetPath = p.Path
'End Function


Public Function GetVisits() As Kvp
    Set GetVisits = p.Visits
End Function

