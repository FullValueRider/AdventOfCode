VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepairDroid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Public Enum Way

    North = 1
    Sount = 2
    West = 3
    East = 4
    
End Enum

Public Enum MoveEvent

    Wall = 0
    Moved = 1
    Oxygen = 2
    
End Enum

Private Type State

    Computer                        As IntComputer
    XLoc                            As Long
    YLoc                           As Long
    Board                           As ExcelBoard
    
End Type

Private s                           As State


Public Function Debutante(ByVal ipX As Long, ByVal ipY As Long) As RepairDroid

        With New RepairDroid
        
        
            Set Debutante = .Self(ipX, ipY)
        
        End With
         
        Set s.Computer = New Computer
        s.computer.
        Set s.Board = New ExcelBoard
        
        s.Board.PlaceTile s.XLoc, s.YLoc, "D"
        
        
End Function


Public Function Self(ByVal ipX As Long, ipY As Long) As RepairDroid

    p.XLoc = ipX
    p.YLoc = ipY
    
End Function


Public Property Get Command(ByVal ipDirection As Direction)

    Dim myInput As Kvp: Set myInput = New Kvp
    s.Comp.Run
End Property

Public Property Get Program(ByVal ipProgram As Kvp)

    Set s.Computer.Program = ipProgram.Clone
    

End Property



