VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpWireAsLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'KvpWireAsLinks(Index:Long by Link:String("U32" ETC)

Private Type State
    'Key:long by Link:string: "L88"
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

Public Property Let Item(ByVal Key As Long, ByVal Value As String)

    s.HostKvp.Item(Key) = Value
    
End Property


Public Sub AddByIndex(ByVal Value As String)
    s.HostKvp.AddByIndex Value
End Sub

Public Sub AddByIndexFromArray(ByVal Value As Variant)
    s.HostKvp.AddByIndexFromArray Value
End Sub


Public Function Keys() As Variant
    Keys = s.HostKvp.GetKeys
End Function

Public Function Values() As Variant
    Values = s.HostKvp.GetValues
End Function

'@Enumerator
Public Property Get NewEnum() As Variant
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
     Set NewEnum = s.HostKvp.KvpEnum
End Property
