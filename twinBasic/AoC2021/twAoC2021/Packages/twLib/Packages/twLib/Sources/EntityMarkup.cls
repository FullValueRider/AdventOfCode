VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EntityMarkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule AssignedByValParameter
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder("Lib.Fmt")

Const MY_NAME                   As String = "EntityMarkup"
Private Type Properties

    Separator                   As String
    Left                        As String
    Right                       As String
    
End Type

Private p                       As Properties

Public Function Deb _
( _
    Optional ByVal ipLeft As String = char.twLBracket, _
    Optional ByVal ipSeparator As String = char.twComma, _
    Optional ByVal ipRight As String = char.twrbracket _
) As EntityMarkup

    With New EntityMarkup
    
        Return .ConstructInstance(ipLeft, ipSeparator, ipRight)
        
    End With
    
End Function

Public Function ConstructInstance _
( _
    ByVal ipLeft As String, _
    ByVal ipSeparator As String, _
    ByVal ipRight As String _
) As EntityMarkup
    
    p.Separator = ipSeparator
    p.Left = ipLeft
    p.Right = ipRight
    
    Return Me
    
End Function

Public Property Get Separator() As String
   Return p.Separator
End Property

Public Property Get Left() As String
    Return p.Left
End Property

Public Property Get Right() As String
    Return p.Right
End Property

Public Function Clone() As EntityMarkup
    Return EntityMarkup.Deb(p.Left, p.Separator, p.Right)
End Function

Public Function ToString() As String
    Return p.Left & "," & p.Separator & "," & p.Right
End Function

Public Property Get Name() As String
    Return MY_NAME
End Property

Public Function ToArray() As Variant
    Return Array(p.Left, p.Separator, p.Right)
End Function