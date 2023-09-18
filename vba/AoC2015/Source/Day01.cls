VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Days")


Option Explicit

Private Const TODAY             As String = "\Day01.txt"


Private Type State

    Data                            As SeqA
    Direction                       As Collection
    
End Type

Private s                           As State

Const UP                            As String = "("
Const DOWN                          As String = ")"
    

Public Sub Execute()
    Part01
    Part02
End Sub

Private Sub Part01()

        Initialise
        
        Dim myResult As Long
        myResult = s.Data.Count(UP) - s.Data.Count(DOWN)
        fmt.Dbg "The answer to Day {0} Part 01 is 74.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult
        
    End Sub

    Private Sub Part02()

        Initialise
       
        Dim myFloor As Long
        myFloor = 0
        Dim myIndex As Long
        Dim myResult As Long
        For myIndex = s.Data.FirstIndex To s.Data.LastIndex
        
            myFloor = myFloor + IIf(s.Data.Item(myIndex) = UP, 1, -1)
           
            If myFloor < 0 Then
                
                myResult = myIndex
                Exit For
                
            End If
            
        Next
        
        fmt.Dbg "The answer to Day {0} part 2 is 1795.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult

    End Sub

    Private Sub Initialise()
    
        Set s.Data = SeqA(Filer.GetFileAsString(AoCRawData & Year & TODAY))
        
'        Set s.Direction = New Collection
'        s.Direction.Add -1, DOWN
'        s.Direction.Add 1, UP
        
    End Sub
