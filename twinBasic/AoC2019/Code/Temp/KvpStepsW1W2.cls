VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpStepsW1W2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'kvp(Index by XYSteps(StepsW1:long by StepsW2:Long)"
Private Type State
    ' key:objPoint by kvp of key:long by TotalSteps:long
    HostKvp As KvpOD
End Type

Private s As State

Private Sub Class_Initialize()
    Set s.HostKvp = New KvpOD
End Sub


'@DefaultMember
Public Property Get Item(ByVal Key As Long) As StepsW1W2
Attribute Item.VB_UserMemId = 0
    Set Item = s.HostKvp.Item(Key)
End Property

Public Property Let Item(ByVal Key As Long, ByVal Value As StepsW1W2)
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
'    Static myKeys As Variant
'    If IsArray(myKeys) Then myKeys = s.HostKvp.GetKeys
'
'    Static myNext As Long
'    If myNext > UBound(myKeys) Then
'        myNext = 0
'        Set NewEnum = Null
'
'    Else
'        Set NewEnum = s.HostKvp.Item(myKeys(myNext))
'    End If
End Function
