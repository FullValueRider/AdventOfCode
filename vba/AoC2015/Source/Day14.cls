VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Days")
Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Private Const TODAY             As String = "\Day14.txt"

Private Type State

    Data                            As seqC
    Reindeer                        As KvpC
    
End Type

Private s                           As State


Public Sub Execute()
    Part01
    Part02
End Sub

Private Sub Part01()

    Initialise
   
    Dim myMaxDistance As Long: myMaxDistance = globals.MinLong
    Dim myDistance As Long: myDistance = 0
    Dim myReindeers As Iteritems: Set myReindeers = Iteritems(s.Reindeer)
    Do
        myDistance = myReindeers.curitem(0).Distance(2503)
        If myDistance > myMaxDistance Then
             myMaxDistance = myDistance
        End If
    Loop While myReindeers.MoveNext
    
    Dim myResult As Long: myResult = myMaxDistance
    
    fmt.dbg "The answer to Day {0} part 1 is {1}.  Found is 2640", VBA.Mid$(TODAY, 5, 2), myResult
    
End Sub

Private Sub Part02()

    Initialise
    
    Dim myScoring As KvpC: Set myScoring = KvpC.Deb.AddPairs(s.Reindeer.Keys, seqC.Deb.Fill(0, s.Reindeer.Count))
    Dim myDistances As KvpC: Set myDistances = KvpC.Deb.AddPairs(s.Reindeer.Keys, seqC.Deb.Fill(0, s.Reindeer.Count))
    Dim myCurrentTime As Long
    For myCurrentTime = 1 To 2503
    
        'Dim myMaxDistance As Long: myMaxDistance = globals.MinLong
        Dim myReindeers As Iteritems: Set myReindeers = Iteritems(s.Reindeer)
        Do
           ' Set myDistances = s.Reindeer.MapIt(mpCallByName.Deb("Distance", myCurrentTime, vbMethod))
                myDistances.Item(myReindeers.curkey(0)) = myReindeers.curitem(0).Distance(myCurrentTime)
        Loop While myReindeers.MoveNext
    '
        Dim myCurrentMaxDistance As Long: myCurrentMaxDistance = myDistances.ReduceIt(rdMaxNum.Deb)
        
        Dim myCurrentWinningReindeer As Iteritems: Set myCurrentWinningReindeer = Iteritems(myDistances.FilterIt(cmpEQ(myCurrentMaxDistance)))
        Do
            myScoring.Item(myCurrentWinningReindeer.curkey(0)) = myScoring.Item(myCurrentWinningReindeer.curkey(0)) + 1
        Loop While myCurrentWinningReindeer.MoveNext
    Next
    
    Dim myResult As Long: myResult = myScoring.ItemsAsSeq.ReduceIt(rdMaxNum)
    
    fmt.dbg "The answer to Day {0} part 2 is 1102.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult

End Sub

Private Sub Initialise()

    Set s.Data = seqC(Filer.GetFileAsArrayOfStrings(AoCRawData & Year & TODAY)).MapIt(mpExecDeb(Reindeer))
  
    Set s.Reindeer = KvpC.Deb
    
    Dim myItems As Iteritems: Set myItems = Iteritems(s.Data)
    Do
        Dim myReindeer As Reindeer: Set myReindeer = myItems.curitem(0)
        s.Reindeer.Add myReindeer.Name, myReindeer
    Loop While myItems.MoveNext
    
End Sub
