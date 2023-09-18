VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OreEater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Private Type Property

    Name                                As String
    Count                               As Long
    Chemical                            As Long
    Ore                                 As Long
    
End Type


Public Function Debutante(ByVal ipDescription As String)

    With New Oreeater
    
    
        Set Debutante = .Self(ipDescription)
    
    End With
    
End Function

    
Public Function Self(ByVal ipCount As Long, ByVal ipchemical As Long, ByVal ipOre As Long)

    p.Count = ipCount
    p.Chemical = ipchemical
    p.Ore = ipOre
    
    Dim myArray As Variant
    myArray = Split(myItem, "=>")
    
    p.Name = Trim(Split(myArray(0), " ")(0))
    p.Count = CLng(Trim(Split(myArray(0), " ")(1)))
    p.Chemical = CLng(Trim(Split(myArray(1), " ")(0)))
    p.Ore = CLng(Trim(Split(myArray(1), " ")(1)))
            
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
