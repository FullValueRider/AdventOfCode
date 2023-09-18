VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OreConsumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Exposed
'@PredeclaredId
Option Explicit

Private Type Properties

    Name                                As String
    Count                               As Long
    Chemical                            As Long
    Ore                                 As Long
    
End Type

Private p                               As Properties


Public Function Debutante(ByVal ipDescription As String) As OreConsumer

    With New OreConsumer
    
        Set Debutante = .Self(ipDescription)
    
    End With
    
End Function

    
Public Function Self(ByVal ipDescription As String) As OreConsumer

    Dim myArray As Variant
    myArray = Split(ipDescription, "=>")
    
    p.Name = Trim$(Split(myArray(0), " ")(0))
    p.Count = CLng(Trim$(Split(myArray(0), " ")(1)))
    p.Chemical = CLng(Trim$(Split(myArray(1), " ")(0)))
    p.Ore = CLng(Trim$(Split(myArray(1), " ")(1)))
    Set Self = Me
    
End Function


Public Property Get Count() As Long
    Count = p.Count
End Property


Public Property Let Count(ByVal ipValue As Long)
    p.Count = ipValue
End Property


Public Property Get Name() As String
    Name = p.Name
End Property
