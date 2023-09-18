VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
Option Explicit

Private Const RAW_DATA_PATH         As String = "C:\Users\slayc\source\repos\VBA\AdventOfCode\2015\Day06RawData.txt"

Private Type State

    Display                                     As BulbDisplay
    RawData                                     As Kvp
    Instructions                                As Kvp
    Lights                                      As Variant

End Type

Private s                                       As State


Private Sub Class_Initialize()

    Set s.RawData = Common.GetRawDataByLine(RAW_DATA_PATH)
    Set s.Instructions = New Kvp
    
    Dim myPair As KVPair
    For Each myPair In s.RawData
    
        Dim myArray As Variant
        myArray = Split(Triage(myPair.Value))
        
        s.Instructions.AddByIndex Array(myArray(0&), Array(CLng(myArray(1&)), CLng(myArray(2&)), CLng(myArray(3)), CLng(myArray(4))))
    
    Next
    
End Sub

Private Function Triage(ByVal ipString As String) As String

     Dim myString As String
     myString = ipString
     myString = Replace(VBA.LCase$(myString), "through ", vbNullString)
     myString = Replace(myString, "turn ", vbNullString)
     myString = Replace(myString, ",", " ")
     
     Dim myLen As Long
    
     Do
     
         myLen = Len(myString)
         myString = Replace(myString, "  ", " ")
         
     Loop Until myLen = Len(myString)
     
     Triage = myString

End Function


Public Sub Part1()

    Dim myArea As Variant
    myArea = Array(0, 0, 999, 999)
    Set s.Display = BulbDisplay.Make(myArea)
    s.Display.SwitchOff myArea
    
    Dim myPair As KVPair
    For Each myPair In s.Instructions
    
        Select Case myPair.Value(0&)

            Case "on": s.Display.SwitchOn myPair.Value(1&)

            Case "off":  s.Display.SwitchOff myPair.Value(1&)

            Case "toggle": s.Display.Toggle myPair.Value(1&)

        End Select

    Next
    ' 486323 is too low
    ' 488495 is too low
    Debug.Print "The answer for Day 06 Part 1 is 543903. Found is ", s.Display.LitBulbs(myArea)

End Sub


Public Sub Part2()


    Dim myArea As Variant
    myArea = Array(0, 0, 999, 999)
    Set s.Display = BulbDisplay.Make(myArea)
    s.Display.SwitchOff myArea
    s.Display.UseBrightness = True
    
    Dim myPair As KVPair
    For Each myPair In s.Instructions
    
        Select Case myPair.Value(0&)

            Case "on": s.Display.SwitchOn myPair.Value(1&)

            Case "off":  s.Display.SwitchOff myPair.Value(1&)

            Case "toggle": s.Display.Toggle myPair.Value(1&)

        End Select

    Next
    ' 486323 is too low
    ' 488495 is too low
    Debug.Print "The answer for Day 06 Part 2 is XXXX. Found is ", s.Display.Brightness(myArea)


End Sub
