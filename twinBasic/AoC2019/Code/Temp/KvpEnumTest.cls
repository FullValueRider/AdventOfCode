VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpEnumTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type State
    
    HostKvp As KvpOD
    
End Type

Private s As State

Private Sub Class_Initialize()

    Set s.HostKvp = New KvpOD
    
End Sub

'@DefaultMember
Public Property Get Item(ByVal Key As Long) As String
    Debug.Print "Property Get"
    Item = s.HostKvp.Item(Key)
    
End Property

Public Property Let Item(ByVal Key As Long, ByVal Value As String)
    Debug.Print "Property Let"
    s.HostKvp.Item(Key) = Value
    
End Property


Public Sub AddByIndex(ByVal Value As String)
    s.HostKvp.AddByIndex Value
End Sub

Public Sub AddByIndexFromArray(ByVal Value As Variant)
    s.HostKvp.AddByIndexFromArray Value
End Sub

Public Function GetValuesAsString() As String
    GetValuesAsString = s.HostKvp.GetValuesAsString
End Function

Public Function GetKeysAsString() As String
    GetKeysAsString = s.HostKvp.GetKeysAsString
End Function

Public Function Keys() As Variant
    Keys = s.HostKvp.GetKeys
End Function

Public Function Values() As Variant
    Values = s.HostKvp.GetValues
End Function

'@Enumerator
Public Property Get NewEnum() As Variant
Attribute NewEnum.VB_UserMemId = -4

    'Dim myVar As KVPair
    'Set myVar = s.HostKvp.KvpEnum
     Set NewEnum = s.HostKvp.KvpEnum
    
End Property

