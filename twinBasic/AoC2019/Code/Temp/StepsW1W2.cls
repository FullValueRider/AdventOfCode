VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StepsW1W2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@PredeclaredId

Private Type Properties

    W1Steps As Long
    W2Steps As Long

End Type

Private p                       As Properties

Public Function Make(ByVal ipW1Steps As Long, ByVal ipW2Steps As Long) As StepsW1W2

    With New StepsW1W2
    
        Set Make = .Self(ipW1Steps, ipW2Steps)
    
    End With
    
End Function


Public Function Self(ByVal ipW1Steps As Long, ByVal ipW2Steps As Long) As StepsW1W2

    p.W1Steps = ipW1Steps
    p.W2Steps = ipW2Steps
    Set Self = Me
    
End Function


Public Function ToString() As String
    ToString = Fmt("{ W1Steps:{0}, W2Steps:{1} )", Array(p.W1Steps, p.W2Steps))
End Function


Public Function FromString(ByVal ipString As String) As Point

    'String is of the form "{ W1Steps:{0}, W2Steps:{1} }""
    Dim myW1W2Steps As Variant
    ' Fill with strings
    myW1W2Steps = Split(ipString, ":")
    ' Replace strigs with numbers
    myW1W2Steps(1) = CLng(Split(myW1W2Steps(1), ",")(0))
    myW1W2Steps(2) = CLng(Trim$(0))
    
    Set FromString = Point.Make(myW1W2Steps(1), myW1W2Steps(2))
            
End Function


