VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpTotalSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit


'KvpTotalSteps            Index:Long by TotalSteps:Long
Private Type State
    'Key:long by Link:string: "L88"
    HostKvp As KvpOD
End Type

Private s As State

Private Sub Class_Initialize()

    Set s.HostKvp = New KvpOD
    
End Sub

'@DefaultMember
Public Property Get Item(ByVal Key As Long) As Long
Attribute Item.VB_UserMemId = 0

    Item = s.HostKvp.Item(Key)
    
End Property

Public Property Let Item(ByVal Key As Long, ByVal Value As Long)

    s.HostKvp.Item(Key) = Value
    
End Property


Public Sub AddByIndex(ByVal Value As Long)
    s.HostKvp.AddByIndex Value
End Sub


Public Function Keys() As Variant
    Keys = s.HostKvp.GetKeys
End Function

Public Function Values() As Variant
    Values = s.HostKvp.GetValues
End Function

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Set NewEnum = s.HostKvp.NewEnum
End Function
    

