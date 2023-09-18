VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Days")


Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Private Const TODAY             As String = "\Day02.txt"


Private Type State
    Data                        As SeqA
End Type

Private s                       As State

Public Sub Execute()
    
        Part01
        Part02
        
End Sub

Public Sub Part01()

    Initialise
    
    Dim myResult As Long
    myResult = s.Data.mapIt(mpExecDeb(Box)).mapIt(mpCallByName("WrappingSize", VbMethod)).ReduceIt(rdSum)

    fmt.Dbg "The answer to Day {0} part 1 is {1}.  Found is {2}", VBA.Mid$(TODAY, 5, 2), "1598415", myResult

End Sub


Public Sub Part02()

    Initialise
   
    Dim myResult As Long
    myResult = s.Data.mapIt(mpExecDeb(Box)).mapIt(mpCallByName("RibbonLength", VbMethod)).ReduceIt(rdSum)
    
    fmt.Dbg "The answer to Day {0} part 2 is {1}.  Found is {2}", VBA.Mid$(TODAY, 5, 2), "3812909", myResult
    
End Sub

Private Sub Initialise()
    Dim myData As Variant
    myData = Filer.GetFileAsArrayOfStrings(AoCRawData & Year & TODAY)
    Set s.Data = SeqA(myData)
End Sub
