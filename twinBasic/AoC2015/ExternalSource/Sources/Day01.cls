VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C


Private Const InputData             As String = "Day01.txt"

Private Type State

    Data                            As seq
    
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
    
    Dim myResult As Long = s.Data.CountIt(cmpEQ("(")) - s.Data.CountIt(cmpEQ(")"))
    
    Fmt.Dbg "The answer to Day {0} part 1 is {1}.  Found is {2}", VBA.Mid$(InputData, 4, 2), "74", myResult
    
End Sub

Private Sub Part02()

    Initialise
    Dim myFloor As Long = 0
    
    Dim myMove As IterItems = IterItems(s.Data)
    Do
    
        myFloor += IIf(myMove.Item = UP, 1, -1)
        If myFloor = -1 Then
            Dim myResult As Long = myMove.Index
            Exit Do
        End If
        
    Loop While myMove.MoveNext
    
    Fmt.Dbg "The answer to Day {0} part 1 is {1}.  Found is {2}", VBA.Mid$(InputData, 4, 2), "1795", myResult

End Sub

Private Sub Initialise()
    Set s.Data = Seq.Deb(Filer.GetFileAsString(AoC2015Data & InputData))
End Sub
