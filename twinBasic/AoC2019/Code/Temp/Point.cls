VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@PredeclaredId

Private Type Properties

    X As Long
    Y As Long

End Type

Private p                               As Properties

Public Function Make(ByVal ipX As Long, ByVal ipY As Long) As Point

    With New Point
    
        Set Make = .Self(ipX, ipY)
    
    End With
    
End Function


Public Function Self(ByVal ipX As Long, ByVal ipY As Long) As Point

    p.X = ipX
    p.Y = ipY
    Set Self = Me
    
End Function


Public Function ToString() As String
    ToString = Fmt("{ X:{0}, Y:{1} )", Array(p.X, p.Y))
End Function


Public Function FromString(ByVal ipString As String) As Point

    'String is of the form "{ X:{0}, Y:{1} }""
    Dim XY As Variant
    ' Fill with strings
    XY = Split(ipString, ":")
    ' Replace strigs with numbers
    XY(0) = CLng(Split(XY(1), ",")(0))
    XY(1) = CLng(Split(XY(1), " ")(0))
    
    Set FromString = Point.Make(XY(0), XY(1))
            
End Function
