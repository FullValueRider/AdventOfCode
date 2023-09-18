VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Asteroid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
Option Explicit

Private Type Properties

    Bearing                         As Double
    Distance                        As Double
    X                               As Long
    Y                               As Long

End Type

Private p                           As Properties

'Private Type State
'
'
'End Type

'Private s As State


Public Function Debut(ByVal ipX As Long, ByVal ipY As Long, Optional ByVal ipBearing As Single = 0, Optional ByVal ipDistance As Single = 0) As Asteroid

    With New Asteroid
    
        Set Debut = .Self(ipX, ipY, ipBearing, ByVal ipDistance)
    
    End With
    
End Function


Public Function Self(ByVal ipX As Long, ByVal ipY As Long, ByVal ipBearing As Single, ByVal ipDistance As Single) As Asteroid

    With p
    
        .X = ipX
        .Y = ipY
        .Bearing = ipBearing
        .Distance = ipDistance
        
    End With
        
    Set Self = Me

End Function


Public Function XCoordinate() As Long
    XCoordinate = p.X
End Function


Public Function YCoordinate() As Long
    YCoordinate = p.Y
End Function


Public Function Distance() As Single
    Distance = p.Distance
End Function


Public Function Bearing() As Single
    Bearing = p.Bearing
End Function


Public Sub UpdateBearingAndDistanceFromOrigin(ByRef ipOrigin As Asteroid)

    p.Distance = Common.DistanceFromOrigin(ipOrigin.XCoordinate, ipOrigin.YCoordinate, Me.XCoordinate, Me.YCoordinate)
    p.Bearing = Common.BearingFromOrigin(ipOrigin.XCoordinate, ipOrigin.YCoordinate, Me.XCoordinate, Me.YCoordinate)
    
End Sub
