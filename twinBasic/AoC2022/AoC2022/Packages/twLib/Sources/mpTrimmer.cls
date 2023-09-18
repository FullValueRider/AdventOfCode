VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpTrimmer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const MY_TYPENAME       As String = "mpTrimer"

Private Type State

    TrimChars                   As Variant
        
End Type

Private s                       As State

Implements IMapper

' Private Function IMapper_ExecMapper(ByVal ipHostItem As Variant) As Variant
'     IMapper_ExecMapper = ExecMapper(ipHostItem)
' End Function

Private Sub Class_Initialize()
    s.TrimChars = Array(" ")
End Sub

Public Function ExecMapper(ByRef ipHostItem As Variant) As Variant Implements IMapper.ExecMapper

    If VBA.VarType(ipHostItem) <> vbString Then
        ExecMapper = ipHostItem
        Exit Function
        
    End If
    
    ExecMapper = Strs.Trimmer(ipHostItem, s.TrimChars)
        
End Function

Public Function Deb(ParamArray ipArgs() As Variant) As mpTrimmer
    With New mpTrimmer
        Return .ConstructInstance(CVar(ipArgs))
    End With
End Function

Friend Function ConstructInstance(ByRef ipArgs As Variant) As mpTrimmer
    s.TrimChars = ipArgs
    Return Me
End Function

Public Function TypeName() As String Implements Imapper.TypeName
    Return MY_TYPENAME
End Function
