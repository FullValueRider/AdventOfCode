VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpWires"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type State
    ' Key:Long by Links:String  "L88,R67,U89 etc"
    HostKvp As KvpOD
End Type

Private s As State

Private Sub Class_Initialize()
    Set s.HostKvp = New KvpOD
End Sub

'@DefaultMember
Public Property Get Item(ByVal Key As Long) As String
Attribute Item.VB_UserMemId = 0
    Item = s.HostKvp.Item(Key)
End Property
   

Public Property Set Item(ByVal Key As Long, ByVal Value As KvpOD)
   s.HostKvp.Item(Key) = Value
End Property

Public Property Set Kvp(ByVal Value As KvpOD)
    Set s.HostKvp = Value
End Property

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = s.HostKvp.NewEnum

End Function
