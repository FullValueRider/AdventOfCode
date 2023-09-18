VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@PredeclaredId
Private Type State

    RawData                         As Kvp
    Boxes                           As Kvp
End Type

Private s                       As State


Public Sub Part1()

    Set s.RawData = Common.GetRawDataByLine("C:\Users\slayc\source\repos\VBA\AdventOfCode\2015\Day02RawData.txt")
    Set s.Boxes = New Kvp
    
    Dim myItem As Variant
    For Each myItem In s.RawData
    
        s.Boxes.AddByIndex Box.Make(myItem)
        
    Next

    
    Dim myTotalArea As Long
    myTotalArea = 0
    Dim myPair As KVPair
    For Each myPair In s.Boxes
        
        Dim myBox As Box
        Set myBox = myPair.Value
        myTotalArea = myTotalArea + myBox.WrappingSize
        
    Next
    
    Debug.Print "The answer to Day 2 Part 1 is 1606483. Found is ", myTotalArea
    
End Sub


Public Sub Part2()

    
    Dim myTotalRibbonLength As Long
    myTotalRibbonLength = 0
    Dim myPair As KVPair
    For Each myPair In s.Boxes
        
        Dim myBox As Box
        Set myBox = myPair.Value
        myTotalRibbonLength = myTotalRibbonLength + myBox.RibbonLength
        
    Next
    
    Debug.Print "The answer to Day 2 Part 2 is 3842356. Found is ", myTotalRibbonLength
    
End Sub
