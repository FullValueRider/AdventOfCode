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

Public Enum Go

    North
    Up
    Forward
    
    South
    Down
    Backward
    
    East
    Right
    
    West
    Left
    
    NW
    NE
    SE
    SW
    
End Enum

Private Type Properties

    X                                As Long
    Y                                As Long
    
End Type



Private p As Properties

Private Type State

	X                                 As Long
    Y                                 As Long
    
End Type

Private s                               As state

'@DefaultMember
Public Function Deb(ByVal ipX As Long, ByVal ipY As Long) As Coord
    With New Coord
        Return .ConstructInstance(CLng(ipX), CLng(ipY))
    End With
End Function


Public Function Deb(ByVal ipXYCoord As String) As Coord
    
    Dim myCoord As Variant = Split(ipXYCoord, Char.twComma)
    Return Deb(CLng(myCoord(0)), CLng(myCoord(1)))

End Function

Public Function Deb(ByVal ipKVPair As KVPair) As Coord

    Return Deb(CLng(ipKVPair.Key(0)), CLng(ipKVPair.Item(ipKVPair.Key(0))))
    
End Function


Public Function ConstructInstance(ByVal ipX As Long, ByVal ipY As Long) As Coord

    p.X = ipX
    p.Y = ipY
    
    s.X = ipX
    s.Y = ipY
    Return Me
    
End Function

Public Function AtOrigin() As Boolean
    Return (s.X = p.X) And (p.Y = s.Y)
End Function

Public Function Reset() As Coord
    p.X = s.X
    p.Y = s.Y
    Return Me
End Function

Public Property Get X() As Long
    X = p.X
End Property

Public Property Let X(ByVal ipValue As Long)
    p.X = ipValue
End Property

Public Property Get Y() As Long
    Y = p.Y
End Property

Public Property Let Y(ByVal ipValue As Long)
    p.Y = ipValue
End Property


Public Property Get ToString() As String
    Return Fmt.Text("{0},{1}", p.X, p.Y)
End Property


Public Function Clone() As Coord
    Return Deb(s.X, s.Y).MoveTo(p.X, p.Y)
End Function

Public Function MoveTo(ByVal ipX As Long, ipY As Long) As Coord
    p.X = ipX
    p.Y = ipY
    Return Me
End Function

Public Function Move(ByVal ipGo As Go, Optional ByVal ipdistance As Long = 1) As Coord

    ' vertical component
    Select Case ipGo
    
        Case North, Up, Forward, NE, NW:        p.Y += ipdistance
        Case South, Down, Backward, SE, SW:     p.Y -= ipdistance
            
    End Select
    
    'Horizontal component
    Select Case ipGo
    
        Case East, Right, SE, NE: :           p.X += ipdistance
        Case West, Left, SW, NW:              p.X -= ipdistance
        
    End Select
    
    Return Me
    
End Function


Public Function Manhatten() As Long
    Return Abs(p.X - s.X) + Abs(p.Y - s.Y)
End Function

Public Function Heading() As Double
    Debug.Print "Need to fix Heading"
    Dim myDx As Double
    myDx = CDbl(p.X - s.X)
    
    Dim myDy As Double
    myDy = CDbl(p.Y - s.Y)
    
    Dim myAngle As Double
    myAngle = Maths.ATan2(myDy, myDx) * Maths.RadiansToDegrees
    
    If myAngle < 0 Then
    
      '  myAngle = myAngle + p.HeadingDatum + 360   ' This is simular to doing
                                    ' 360 Math.Atan2(y1 - y2, x1 - x2) * (180 / Math.PI)
                                    
    End If
    
    Heading = myAngle
   
End Function