VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@PredeclaredId
'@IgnoreModule

Private Type Keys

    X                               As Variant
    Y                               As Variant
    Item                            As Variant
    Values                          As Variant

End Type

Private Type State

    Map                             As KvpOD
    InitialKeys                     As Keys
    CurrentKeys                     As Keys

End Type

Private s                           As State


Public Function Create _
( _
    ByVal ipXKeys As KvpOD, _
    ByVal ipYKeys As KvpOD, _
    Optional ByVal ipItemKeys As KvpOD, _
    Optional ByVal ipItemValues As KvpOD _
 ) As KvpMap
 
    With New KvpMap
    
        Set Create = .Debutante(ipXKeys, ipYKeys, ipItemKeys, ipItemValues)
    
    End With
    
End Function


Public Function Debutante _
( _
    ByVal ipXKeys As KvpOD, _
    ByVal ipYKeys As KvpOD, _
    Optional ByVal ipItemKeys As KvpOD, _
    Optional ByVal ipItemValues As KvpOD _
 ) As KvpMap

    If ipItemKeys.Count < ipItemValues.Count Then
    
        Debug.Print "The number of value items should not exceed the number of value Keys"
        End
        
    End If

    With s.InitialKeys
    
        Set .X = ipXKeys.Clone
        Set .Y = ipYKeys.Clone
        Set .Item = ipItemKeys.Clone
        Set .Values = ipItemValues.Clone
        
    End With
    
    With s.CurrentKeys
    
        Set .X = ipXKeys.Clone
        Set .Y = ipYKeys.Clone
        Set .Item = ipItemKeys.Clone
        Set .Values = ipItemValues.Clone
        
    End With
    
    Set s.Map = New KvpOD
    Dim myY As Variant
    For Each myY In s.InitialKeys.Y
        
        s.Map.AddByKey myY, New KvpOD
        
        Dim myX As Variant
        For Each myX In s.InitialKeys.X
        
            'AddKey s.Map.Item(myY), myX
            s.Map.Item(myY).AddByKey myX, New KvpOD
        
            ' Now add the values
            
            Dim myItemKeys As Variant: myItemKeys = s.InitialKeys.Item.GetKeys
            Dim myValueKeys As Variant: myValueKeys = s.InitialKeys.Item.GetValues
            Dim myIndex As Long
            For myIndex = LBound(myItemKeys) To UBound(myItemKeys)
            
                With s.Map.Item(myY).Item(myX)
                
                    .AddByKey _
                        s.InitialKeys.Item.Item(myItemKeys(myIndex)), _
                        s.InitialKeys.Item.Item(myValueKeys(myIndex))
                
                End With
            
            Next
            
            
        Next
        
    Next
    
    Set Debutante = Me
    
End Function

Public Function NewEnum() As IUnknown
    Set NewEnum = s.Map.NewEnum
End Function

'Public Sub AddKey(ByRef ipKvp As Kvp, ByVal ipKey As Variant)
'    Debug.Print TypeName(ipKey), VarType(ipKey)
'    Select Case VarType(ipKey)
'
'        Case vbLong, vbInteger:   ipKvp.AddByKey CLng(ipKey), New KvpOD
'        Case vbSingle:  ipKvp.AddByKey CSng(ipKey), New KvpOD
'        Case vbString:  ipKvp.AddByKey CStr(ipKey), New KvpOD
'
'        Case Else
'
'            Debug.Print "Type not supperted", TypeName(ipKey)
'
'    End Select
'
'End Sub


'Public Sub Show(ByVal ipKey As Variant)
'
'End Sub
'Public Sub AddLayer(ByVal ipKey As Variant, Optional ByVal ipValue As Variant = " ")
'
'End Sub
