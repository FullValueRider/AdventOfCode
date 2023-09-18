VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Classes")
'@PredeclaredId

'@IgnoreModule
Option Explicit

Private Type State

    Host               As kvpair
    
End Type

Private s               As State

Private Sub Class_Initialize()
    
    Set s.Host = KVPair.Deb
    
End Sub



Public Function Deb(ParamArray ipArgs() As Variant) As Tuple

    With New Tuple
    
        Set Deb = .ReadyToUseInstance(CVar(ipArgs))
    
    End With
    
End Function

Public Function ReadyToUseInstance(ByVal ipArgs As Variant) As Tuple

    s.Host = ipArgs
    
    Set ReadyToUseInstance = Me
    
End Function

Public Property Get Item(ByVal ipIndex As Long) As Variant

    
    Item = s.Host(ipIndex)
        
End Property
