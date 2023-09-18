VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day15TrackPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredelaredId
Option Explicit

Private Type Properties

    Steps                       As Long
    myType                      As TileType
    NType                       As TileType
    SType                       As TileType
    EType                       As TileType
    WType                       As TileType
    X                           As Long
    Y                           As Long
    
End Type

Public Function Debutanate(ByVal ipX As Long, ByVal ipY As Long, ByVal ipType As Long) As Day15TrackPoint


    With New Day15TrackPoint
    
        Set Debutante = .Self(ipX, ipY, ipType)
        
    End With
    
End Function

Public Function Self(ByVal ipX As Long, ByVal ipY As Long, ByVal ipType As Long) As Day15TrackPoint

    p.X = ipX
    p.Y = ipY
    p.myType = ipType
    p.NType = -1
    p.SType = -1
    p.WType = -1
    p.EType = -1
    
    Set Self = Me
    
    
End Function


Public Function GetTrackPointKey()
    GetTrackPointKey = "(" & CStr(s.Tp.X) & "," & CStr(s.Tp.Y) & ")"
End Function


Private Function GetNeighbourTrackKey(ByVal ipWay As Way) As String

    Dim myX As Long
    myX = p.X
    
    Dim myY As Long
    myY = p.Y

    Select Case ipWay
    
        Case Way.North: myY = myY + 1
        Case Way.South: myY = myY - 1
        Case Way.west: myX = myX + 1
        Case Way.east: myX = myX - 1
    
    End Select
    
    GetNeighbourTrackKey = "(" & CStr(myX) & "," & CStr(myY) & ")"
    
End Function
