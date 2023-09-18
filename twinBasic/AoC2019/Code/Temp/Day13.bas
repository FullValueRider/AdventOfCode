Attribute VB_Name = "Day13"
Option Explicit

#If VBA7 Then
    Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Declare Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
#End If




Public Enum Tile

    IsEmpty = 0
    IsWall
    IsBlock
    IsHPaddle
    IsBall
    

End Enum



'Public Piece(IsEmpty To IsBall)                     As String

'Public Sub Day13Part1()
'
'    Dim myGameBoard As ExcelBoard: Set myGameBoard = New ExcelBoard
'    InitialisePieces
'    Dim myComp As IntComputer: Set myComp = New IntComputer
'    Dim myProgram As Kvp
'    Set myProgram = GetDay13Program
'    Set myComp.Program = myProgram
'    myComp.OutputMode = HaltOnOutput
'    Dim myLayout As KvpOD: Set myLayout = New KvpOD
'
'    Dim myX As Long
'    Dim myY As Long
'    Dim myPiece As Tile
'
'    Do
'
'        myComp.Run
'        If myComp.RunHasCompleted Then Exit Do
'        myX = myComp.GetOutput.GetFirst.Value
'        myComp.Run
'        myY = myComp.GetOutput.GetFirst.Value
'        myComp.Run
'        myPiece = myComp.GetOutput.GetFirst.Value
'
'        If myLayout.LacksKey(XYCoord(myX, myY)) Then
'
'            myLayout.AddByKey XYCoord(myX, myY), myPiece
'
'        Else
'
'            myLayout.SetItem XYCoord(myX, myY), myPiece
'
'        End If
'
'        myGameBoard.PlaceTile myX, myY, Piece(myPiece)
'
'    Loop Until myComp.RunHasCompleted
'
'
'    Dim myRankKvp As Kvp
'    Set myRankKvp = myLayout.Rank
'    Debug.Print myRankKvp.GetKeysAsString
'    Debug.Print myRankKvp.GetValuesAsString
'    Debug.Print "The answer to Day 12 Part 1 is "; myRankKvp.Item(CStr(Piece(IsBlock)))
'
'End Sub
'Public Sub Day13Part2()
'
'    Dim myGameBoard As ExcelBoard
'    InitialisePieces
'    Dim myComp As IntComputer: Set myComp = New IntComputer
'    Dim myProgram As Kvp
'    Set myProgram = GetDay13Program
'    Set myComp.Program = myProgram
'    myComp.OutputMode = HaltOnOutput
'    myProgram.SetItem 0&, 2
'    myComp.OutputMode = HaltOnOutput
'    Dim myLayout As KvpOD: Set myLayout = New KvpOD
'
'    Dim myX As Long
'    Dim myY As Long
'    Dim myPiece As Tile
'
'    Do
'        'If GetKeyState(VK_LEFT) = True Then Set myInput = MakeKvp(-1)
'        myComp.Run
'        If myComp.RunHasCompleted Then Exit Do
'        myX = myComp.GetOutput.GetFirst.Value
'        myComp.Run
'        myY = myComp.GetOutput.GetFirst.Value
'        myComp.Run
'        myPiece = myComp.GetOutput.GetFirst.Value
'
'        If myX = -1 And myY = 0 Then
'
'            myGameBoard.PlaceTile myX, myY, Chr$(32 + myPiece)
'
'        End If
'
'
'        If myLayout.LacksKey(XYCoord(myX, myY)) Then
'
'            myLayout.AddByKey XYCoord(myX, myY), myPiece
'
'        Else
'
'            myLayout.SetItem XYCoord(myX, myY), myPiece
'
'        End If
'
'    Loop Until myComp.RunHasCompleted
'
'End Sub

Public Function XYCoord(ByVal ipX As Long, ByVal ipY As Long) As String
    XYCoord = "(" & CStr(ipX) & "," & CStr(ipY) & ")"
End Function


Public Function GetDay13Program() As KvpOD

    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day13\Day13Program.txt", ForReading)
    Dim myProgram As String:  myProgram = myfile.ReadAll
    myfile.Close
    Dim myKvp As KvpOD
    Set myKvp = MakeProgram(Split(myProgram, ","))
    Set GetDay13Program = myKvp
    
End Function

'Public Sub InitialisePieces()
'
'    Piece(IsEmpty) = " "
'    Piece(IsWall) = "="
'    Piece(IsBlock) = "-"
'    Piece(IsHPaddle) = "_"
'    Piece(IsBall) = "o"
'
'End Sub

'Public Sub GetTile(ByVal ipTile As Tile)
'
'    Dim myTile As String
'    Select Case ipTile
'
'        Case IsEmpty: myTile = Piece(IsEmpty)
'        Case IsWall: myTile = Piece(IsWall)
'        Case IsBlock: myTile = Piece(IsBlock)
'        Case IsHPaddle: myTile = Piece(IsHPaddle)
'        Case IsBall: myTile = Piece(IsBall)
'        Case Else
'
'            Debug.Print "ShowTilePositionInExcel: Unknown tile type ",
'
'    End Select
'
'   ' s.Piece.PlaceTile ipX, ipy, myTile
'
'End Sub

