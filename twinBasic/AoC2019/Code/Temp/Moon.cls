VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Moon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Private Type Properties

    XPos                        As Long
    YPos                        As Long
    ZPos                        As Long
    XVel                        As Long
    YVel                        As Long
    ZVel                        As Long
    
End Type

Private p                       As Properties


Public Function Debutante _
( _
    ByVal ipXPos As Long, _
    ByVal ipYPos As Long, _
    ByVal ipZPos As Long, _
    Optional ByVal ipXVel As Long = 0, _
    Optional ByVal ipYVel As Long = 0, _
    Optional ByVal ipZVel As Long = 0 _
) As Moon


    With New Moon
    
        Set Debutante = .Self(ipXPos, ipYPos, ipZPos, ipXVel, ipYVel, ipZVel)
        
    End With
    
End Function


Public Function Self _
( _
    ByVal ipXPos As Long, _
    ByVal ipYPos As Long, _
    ByVal ipZPos As Long, _
    ByVal ipXVel As Long, _
    ByVal ipYVel As Long, _
    ByVal ipZVel As Long _
) As Moon

    With p
    
        .XPos = ipXPos
        .YPos = ipYPos
        .ZPos = ipZPos
        .XVel = ipXVel
        .YVel = ipYVel
        .ZVel = ipZVel

    End With
    
    Set Self = Me

End Function

Public Property Get XPos() As Long
    XPos = p.XPos
End Property

Public Property Let XPos(ByVal ipValue As Long)
    p.XPos = ipValue
End Property


Public Property Get YPos() As Long
    YPos = p.YPos
End Property

Public Property Let YPos(ByVal ipValue As Long)
    p.YPos = ipValue
End Property


Public Property Get ZPos() As Long
    ZPos = p.ZPos
End Property

Public Property Let ZPos(ByVal ipValue As Long)
    p.ZPos = ipValue
End Property


Public Property Get XVel() As Long
    XVel = p.XVel
End Property

Public Property Let XVel(ByVal ipValue As Long)
    p.XVel = ipValue
End Property


Public Property Get YVel() As Long
    YVel = p.YVel
End Property

Public Property Let YVel(ByVal ipValue As Long)
    p.YVel = ipValue
End Property


Public Property Get ZVel() As Long
    ZVel = p.ZVel
End Property

Public Property Let ZVel(ByVal ipValue As Long)
    p.ZVel = ipValue
End Property


Public Sub ApplyGravity(ByRef ipMoon As Moon)
    
    If p.XPos < ipMoon.XPos Then
    
        p.XVel = p.XVel + 1
        ipMoon.XVel = ipMoon.XVel - 1
    
    ElseIf p.XPos > ipMoon.XPos Then
    
        p.XVel = p.XVel - 1
        ipMoon.XVel = ipMoon.XVel + 1
        
    End If
    
    If p.YPos < ipMoon.YPos Then
    
        p.YVel = p.YVel + 1
        ipMoon.YVel = ipMoon.YVel - 1
    
    ElseIf p.YPos > ipMoon.YPos Then
    
        p.YVel = p.YVel - 1
        ipMoon.YVel = ipMoon.YVel + 1
        
    End If
    
    If p.ZPos < ipMoon.ZPos Then
    
        p.ZVel = p.ZVel + 1
        ipMoon.ZVel = ipMoon.ZVel - 1
    
    ElseIf p.ZPos > ipMoon.ZPos Then
    
        p.ZVel = p.ZVel - 1
        ipMoon.ZVel = ipMoon.ZVel + 1
        
    End If
    
End Sub


Public Sub ApplyVelocity()

    p.XPos = p.XPos + p.XVel
    p.YPos = p.YPos + p.YVel
    p.ZPos = p.ZPos + p.ZVel
    
End Sub


Public Function PotentialEnergy() As Long
    PotentialEnergy = Abs(p.XPos) + Abs(p.YPos) + Abs(p.ZPos)
End Function


Public Function KineticEnergy() As Long
    KineticEnergy = Abs(p.XVel) + Abs(p.YVel) + Abs(p.ZVel)
End Function


Public Function TotalEnergy() As Long
    TotalEnergy = KineticEnergy * PotentialEnergy
End Function
