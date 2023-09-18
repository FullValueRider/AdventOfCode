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

    Map                             As Kvp
    InitialKeys                     As Keys
    CurrentKeys                     As Keys

End Type

Private s                           As State


Public Function Create _
( _
    ByVal ipXKeys As Kvp, _
    ByVal ipYKeys As Kvp, _
    Optional ByVal ipItemKeys As Kvp, _
    Optional ByVal ipItemValues As Kvp _
 ) As KvpMap
 
    With New KvpMap
    
        Set Create = .Debutante(ipXKeys, ipYKeys, ipItemKeys, ipItemValues)
    
    End With
    
End Function


Public Function Debutante _
( _
    ByVal ipXKeys As Kvp, _
    ByVal ipYKeys As Kvp, _
    Optional ByVal ipItemKeys As Kvp, _
    Optional ByVal ipItemValues As Kvp _
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
    
    Set s.Map = New Kvp
    Dim myY As Variant
    For Each myY In s.InitialKeys.Y
        
        s.Map.AddByKey myY, New Kvp
        
        Dim myX As Variant
        For Each myX In s.InitialKeys.X
        
            'AddKey s.Map.Item(myY), myX
            s.Map.Item(myY).AddByKey myX, New Kvp
        
            ' Now add the values
            
            Dim myItemKeys As Variant: myItemKeys = s.InitialKeys.Item.GetKeys
            Dim myValueKeys As Variant: myValueKeys = s.InitialKeys.Values.GetKeys
            Dim myIndex As Long
            For myIndex = LBound(myItemKeys) To UBound(myItemKeys)
            
                With s.Map.Item(myY).Item(myX)
                
                    .AddByKey _
                        s.InitialKeys.Item(myItemKeys(myIndex)), _
                        s.InitialKeys.Values(myValueKeys(myIndex))
                
                End With
            
            Next
            
            
        Next
        
    Next
    
    Set Debutante = Me
    
End Function



'Public Sub AddKey(ByRef ipKvp As Kvp, ByVal ipKey As Variant)
'    Debug.Print TypeName(ipKey), VarType(ipKey)
'    Select Case VarType(ipKey)
'
'        Case vbLong, vbInteger:   ipKvp.AddByKey CLng(ipKey), New Kvp
'        Case vbSingle:  ipKvp.AddByKey CSng(ipKey), New Kvp
'        Case vbString:  ipKvp.AddByKey CStr(ipKey), New Kvp
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
