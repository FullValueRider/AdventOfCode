VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Number"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@IgnoreModule
'@PredeclaredId

Private Type Properties

    Value As Variant
    
End Type

Public Function Mk(ByVal ipNumber As Variant) As Number

    If Not IsNumeric(ipNumber) Then
    
        With Err
            
            .Clear
            .Raise vbObjectError + 13, "Number.Mk", "The parameter is not a number Type"
            
        End With
        
    End If
        
End Function



Private Function IsNumberType(ByVal ipNumber As Variant) As Boolean

    IsNumberType = False
    If IsObject(ipNumber) Then
    
        'Test for Classes we have created that we wish to be a number
        
    ElseIf IsNumeric(ipNumber) Then
    
        
    ElseIf IsStringy(ipNumber) Then
    
    
    End If
    
End Function
