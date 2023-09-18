VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Day05")

Const InputData           As String = "Day05.txt"

Public Enum VentLineType
    
    vlPoint = 2 ^ 0
    vlHorizontal = 2 ^ 1
    vlVertical = 2 ^ 2
    vlDiagonal = 2 ^ 3
        
End Enum


Private Type State
    
    VentLines                       As Seq ' VentCoord objects
    
End Type

Private s                           As State

Public Sub Execute()
    
    Part01
    Part02
    
End Sub


Public Sub Part01()

    Initialise
    Dim myResult As Long
    myResult = BuildVentMap(s.VentLines, Seq.Deb.AddItems(vlPoint, vlHorizontal, vlVertical)).CountIt(cmpMT.Deb(1))
    
    Fmt.Dbg "The answer to Day 05 part 1 is 7085.  Found is {0}", myResult

End Sub

Public Sub Part02()
    Initialise
    Dim myResult As Long
    myResult = BuildVentMap(s.VentLines, Seq.Deb.AddItems(vlPoint, vlHorizontal, vlVertical, vlDiagonal)).CountIt(cmpMT.Deb(1))
    
    Fmt.Dbg "The answer to Day 05 part 2 is 20271.  Found is {0}", myResult
    
End Sub

Public Sub Initialise()

    Set s.VentLines = _
        Seq _
            .Deb _
            .AddRange(VBA.Split(Filer.GetFileAsString(AoC2021Data & InputData), vbCrLf)) _
            .MapIt(mpMultiReplace.Deb(Array(" ", vbNullString), Array("->", ","))) _
            .MapIt(mpConstruct(VentCoords))
    
End Sub

Public Function BuildVentMap(ByRef ipVentLines As Seq, ByVal ipAllowedLineTypes As Seq) As Hkvp
    
    Dim myVentMap As Hkvp = Hkvp.Deb
    
    Dim myVentLine As Variant
    For Each myVentLine In ipVentLines.Items
        
        If ipAllowedLineTypes.HoldsItem(myVentLine.GetLineType) Then
            
            AddVentMapLine myVentLine, myVentMap
            
        End If
        
    Next
    
    Set BuildVentMap = myVentMap
    
End Function

Public Sub AddVentMapLine(ByVal ipVentLine As VentCoords, ByRef iopVentMap As Hkvp)
    
    Dim mycoords As Seq = ipVentLine.GetCoords
    
    Dim myCoord As Variant ' Is String
    For Each myCoord In mycoords.Items
        
        If iopVentMap.HoldsKey(myCoord) Then
            
            iopVentMap.Item(myCoord) = iopVentMap.Item(myCoord) + 1
        
        Else
            
            iopVentMap.Add myCoord, 1
            
        End If
        
    Next
    
End Sub
    

