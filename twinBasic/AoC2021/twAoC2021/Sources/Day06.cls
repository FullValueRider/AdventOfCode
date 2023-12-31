VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Day02")

Const InputData           As String = "Day06.txt"


Private Type State
    
    Data                            As Seq
    
    
End Type

Private s                           As State


Private Sub class_initialize()
    
    Set s.Data = _
    Seq _
        .Deb _
        .AddRange(VBA.Split(Filer.GetFileAsString(AoC2021Data & InputData), Char.twComma)) _
        .MapIt(mpConvert.Deb(ToLong))
        
End Sub
 
Public Sub Execute()
    
    Part01
    Part02
    
End Sub

Public Sub Part01()

    Dim myResult As LongLong
    myResult = Grow(80, s.Data)
    Fmt.Dbg "The answer to Day {0} part 1 is {1}.  Found is {2}", VBA.Mid$(InputData, 4, 2), "353079", myResult
        
End Sub

Public Sub Part02()

    Dim myResult As LongLong
    myResult = Grow(256, s.Data)
    Fmt.Dbg "The answer to Day {0} part 2 is {1}.  Found is {2}", VBA.Mid$(InputData, 4, 2), "1605400130036", myResult
    
End Sub

Public Function Grow(ByVal ipDays As Long, ByRef ipFishes As Seq) As LongLong

    Dim myHist As Seq = Seq.Deb(Array(0, 0, 0, 0, 0, 0, 0, 0, 0))

    Dim myI As IterItems = IterItems(ipFishes)  'Dim myFishAge As Variant
    Do
        
        myHist.Item(myI.Item + 1) += 1
        
    Loop While myI.MoveNext

    Dim myspawn As LongLong
    Dim myDays As Long
    For myDays = 1 To ipDays
        
        myspawn = myHist.Item(1)
        myHist.RemoveAt 1
        myHist.Add myspawn
        myHist.Item(7) += myspawn
        
    Next

    Return myHist.ReduceIt(rdSum.Deb)

End Function
  
  ' The brute force soilution below works for part 1 at 80 days but when days
  ' is 256 it is very slow and eventally runs out of memory
  ' Public Sub Part02()
      
  '   Dim myLF2 As Seq = Seq.Deb.AddRange(s.Data)
    
  '   Dim myHist As Seq = Seq.Deb.AddRange(Array(0, 0, 0, 0, 0, 0))
    
  '   Dim myItem As Variant
  '   For Each myItem In myLF2
      
  '     myhist.item(myitem) +=1
      
  '   Next
    
  '   Dim myLf As Seq = Seq.Deb.Add(0)
    
  '   Dim myCounter As Long
  '   Do
      
  '     DoEvents
  '     Set myLF = myLF.MapIt(mpDec.Deb(1))
  '     Dim mySpawn As Long = myLF.CountIt(cmpEQ.Deb(-1))
  '     If mySpawn > 0 Then Set myLF = myLF.AddRange(Ranges.GetConstSeries(mySpawn, 8))
  '     Set myLF = myLF.MapIt(mpSetIf.Deb(cmpEQ.Deb(-1), 6))
  '     myCounter = myCounter + 1
  '     Debug.Print myCounter,
      
  '   Loop Until myCounter = 256
    
  '   Dim myResult As LongLong = myLf.Count * myLF2.Count
    
    
  '   Fmt.Dbg "The answer to Day {0} part 2 is {1}.  Found is {2}", VBA.Mid$(InputData, 4, 2), "xxxxxx", myResult
      
  ' End Sub

