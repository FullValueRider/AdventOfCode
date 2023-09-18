VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InputCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Exposed
'@PredeclaredId
'@IgnoreModule
Option Explicit
Private Enum KeyMask

    NewPress = &H1
    Down = &H8000
    Value = &H7FFE
    
End Enum

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Event TestNewTile(ByVal ipCoords As String)
Public Event MoveToKnownTile(ByVal ipCoords As String)

Private Type Properties

    KeyToWay                As KvpD
    
End Type

Private p                   As Properties


'@Ignore WriteOnlyProperty
Public Property Set KeyToWayMap(ByVal ipKeyToWay As KvpOD)
    Set p.KeyToWay = ipKeyToWay
End Property


Public Sub ScanKeyboardForWayToMoveUntilQuit(ByVal ipQuit As Integer)
    
    Do
        
        Dim myKey  As Integer
        myKey = GetAsyncKeyState(vbKeyRight)
        
        Dim MyValue As Long
        MyValue = myKey And KeyMask.Value
        If MyValue = ipQuit Then Exit Do
        
        Dim myDown As Boolean
        myDown = CBool(myKey And KeyMask.Down)
        If myDown Then
        
            Dim myNewPress As Boolean
            myNewPress = CBool(myKey And KeyMask.NewPress)
            If myNewPress Then
                
               ' RaiseEvent WayToMove(p.KeyToWay.Item(myValue))
                
            End If
            
        End If
    
        DoEvents
        
    Loop

End Sub

