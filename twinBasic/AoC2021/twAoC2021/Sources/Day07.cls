VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Day07")


  Const InputData           As String = "Day07.txt"


Private Type State
    
    Data                            As Seq
    
    
End Type

Private s                           As State



Public Sub Execute()
    
    Part01
    Part02
     
End Sub

Public Sub Part01()
  
  Initialise
  Dim myAverageHorizontalPosition As Long = VBA.Round(s.Data.ReduceIt(rdSum) / s.Data.Count, 0)
 
  Dim myHorizontalPosition As Long
  Dim mycost As Long = enums.Preset.Value(MaxLong)
  For myHorizontalPosition = myAverageHorizontalPosition To 0 Step -1
  
    Dim myDist As Seq = s.Data.Clone.MapIt(mpDec(myHorizontalPosition))
    Dim mySumDist As Long = myDist.MapIt(mpMath(Fx.Abs)).ReduceIt(rdSum)
    If mycost < mySumDist Then
        Exit For
    End If
    mycost = mySumDist
    
  Next

  Fmt.Dbg "The answer to Day {0} part 1 is {1}.  Found is {2}", VBA.Mid$(InputData, 4, 2), "343441", mycost
      
End Sub

Public Sub Part02()
    /*
    This method can be refined by calculating the diffs once, 
    then incrementing the diffs over the max min range
    
    */
    Initialise
    Dim myMaxHorizontal As LongLong = s.Data.ReduceIt(rdMax)
    Dim myMinHorizontal As LongLong = s.Data.ReduceIt(rdMin)
    
    Dim myMinFuel As LongLong = enums.Preset.Value(MaxLongLong)
    
    Dim myH As Long
    For myH = CLng(myMinHorizontal) To CLng(myMaxHorizontal)
    
      Dim myFuel As Long = _
          s.Data _
              .Clone _
              .MapIt(mpDec(myH)) _
              .MapIt(mpMath(Fx.Abs)) _
              .MapIt(mpMath(Fx.TriangularNumber)) _
              .ReduceIt(rdSum)
    
        If myFuel < myMinFuel Then
            myMinFuel = myFuel
        Else
            Exit For
        End If
        
    
  Next
  
  Fmt.Dbg "The answer to Day {0} part 2 is {1}.  Found is {2}", VBA.Mid$(InputData, 4, 2), "98925151", myMinFuel
      
End Sub


Private Sub Initialise()
  
    Set s.Data = _
      Seq.Deb _
        .AddRange(VBA.Split(Filer.GetFileAsString(AoC2021Data & InputData), Char.twComma)) _
        .MapIt(mpConvert(ToLong))
      
End Sub


