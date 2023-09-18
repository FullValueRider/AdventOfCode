VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" ()

Const VK_LEFT                   As Long = &H25
Const VK_UP                     As Long = &H26
Const VK_RIGHT                  As Long = &H27
Const VK_DOWN                   As Long = &H28
Const VK_Q                      As Long = &H51

Public Event GoNorth() 'Any object that consumes this event must declare an inputcollector object
                       ' and also include the method preceded with the variable name
                       
Public Event GoSouth()
Public Event GoWest()
Public Event GoEast()
Public Event StopMe()

Public Sub GetKeyPress()

    Do
    
        If GetKeyState(VK_RIGHT And &H8000) Then RaiseEvent GoWest
        If GetKeyState(VK_LEFT And &H8000) Then RaiseEvent GoEast
        If GetKeyState(VK_UP And &H8000) Then RaiseEvent GoNorth
        If GetKeyState(VK_DOWN And &H8000) Then RaiseEvent GoSouth
        If GetKeyState(VK_Q And &H8000) Then RaiseEvent StopMe


End Sub




