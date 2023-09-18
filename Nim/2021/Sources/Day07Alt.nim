import strformat
import sequtils
import strutils
import AoCLib
import math
import system\iterators

const InputData = "Day07.txt"


type 
    State = object
        Data: seq[int]
    
var s: State


proc Initialise()=
    s.Data = readfile(RawDataPath2021 & InputData).split(",").mapIt( it.parseint )


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
    var myMaxHorizontal : int = s.Data.ReduceIt(rdMax)
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





proc Execute*() =
    Part01()
    Part02()