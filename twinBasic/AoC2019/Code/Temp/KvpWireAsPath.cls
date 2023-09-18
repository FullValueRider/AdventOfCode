VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpWireAsPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit
'KvpWireAsPath              Point: by PathTotalSteps(Kvp-> Index:Long by TotalSteps:Long)

Private Type State
    HostKvp                     As KvpOD
End Type

Private s As State

Private Sub Class_Initialize()

    Set s.HostKvp = New KvpOD
    
End Sub



'@DefaultMember
Public Property Get Item(ByVal Key As Point) As KvpTotalSteps
Attribute Item.VB_UserMemId = 0
    
    Set Item = s.HostKvp.Item(Key.ToString)
    
End Property

Public Property Set Item(ByVal Key As Point, ByVal Value As KvpTotalSteps)

    Set s.HostKvp.Item(Key.ToString) = Value
    
End Property


Public Sub AddByKey(ByVal Key As Point, ByVal Value As KvpTotalSteps)
    s.HostKvp.AddByKey Key.ToString, Value
End Sub


Public Function Keys() As Variant
    Keys = s.HostKvp.GetKeys
End Function

Public Function Values() As Variant
    Values = s.HostKvp.GetValues
End Function

Public Function HoldsKey(ByVal Key As Point) As Boolean
    HoldsKey = s.HostKvp.HoldsKey(Key.ToString)
End Function
Public Function LacksKey(ByVal Key As Point) As Boolean
    LacksKey = s.HostKvp.LacksKey(Key.ToString)
End Function

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Set NewEnum = s.HostKvp.[_NewEnum]

End Function
