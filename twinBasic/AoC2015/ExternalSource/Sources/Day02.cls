VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C


Const InputData                                 As String = "Day02.txt"

Private Type State

    Data                                        As seq
    
End Type

Private s                       As State

Public Sub Execute()
	
        Part01
        Part02
        
End Sub

Public Sub Part01()

   Initialise
    
    'Dim myResult As Long = s.Data.MapIt(mpConstruct(Box)).MapIt(mpCallByName("WrappingSize")).ReduceIt(rdSum)
    Dim myResult As Long = 0
    Dim myBox As IterItems = IterItems(s.Data)
    Do
    
        myResult += Box.Deb(myBox.Item).WrappingSize
        
    Loop While myBox.MoveNext
    
    Fmt.Dbg "The answer to Day {0} part 1 is {1}.  Found is {2}", VBA.Mid$(InputData, 4, 2), "1606483", myResult
   
End Sub


Public Sub Part02()

    
    ' Dim myTotalRibbonLength As Long
    ' myTotalRibbonLength = 0
    ' Dim myPair As seqair
    ' For Each myPair In s.Boxes
        
    '     Dim myBox As Box
    '     Set myBox = myPair.Value
    '     myTotalRibbonLength = myTotalRibbonLength + myBox.RibbonLength
        
    ' Next
    
    Fmt.Dbg "The answer to Day {0} part 1 is {1}.  Found is {2}", VBA.Mid$(InputData, 4, 2), "3842356", myResult
    
End Sub

Private Sub Initialise()
    Dim myData As Variant = Filer.GetFileAsArrayOfStrings(AoC2015Data & InputData)
    Set s.Data = Seq.Deb(myData)
End Sub