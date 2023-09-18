import strformat
import sequtils
import strutils
import AoCLib
import math
import system\iterators
import chars
import sequtils
import strutils

const InputData = "Day07.txt"

type
  State = object
    Signals:                        seq[seq[string]]
    Digits:                         seq[seq[string]]

var s = new  State
#var myData= readfile(RawDataPath2021 & InputData).split(chars.twcrlf)
   
    
proc initialise() =

    let myReplaceSpaceBarByBar =(" |", "|")
    let myReplaceBarSpaceByBar = ("| ", "|")
    let myReplaceSpaceSpacesBySpace = ("  ", " ")
    let myPath =RawDataPath2021 & InputData
    var myData : seq[seq[string]] = 
        readfile(myPath).split("\r\n")
        .mapIt(it.multireplace( myreplaceSpaceBarByBar ))
        .mapit(it.multireplace( myReplaceSpaceSpacesBySpace ))
        .mapit(it.multireplace( myReplaceSpaceSpacesBySpace ))
        .mapit(it.split('|'))
    
    s.Signals = myData.mapIt(it[0]).mapIt(it.split(' '))
    s.Digits = myData.mapIt(it[1]).mapIt(it.split(' '))


proc part01() =

    initialise()
    
    var myCount : int = 0
    for myDisplay in s.Digits:
        for myDigit in myDisplay:
            case myDigit.len:
                of 2, 3, 4, 7:
                    myCount += 1   
                else:
                    discard                 
        
    echo fmt"The answer to Day {InputData[4..5])} part 1 is 532.  Found is {myCount}"
    

# The order of populating segments is 
#         3
#       4   1
#         5
#       7   2
#         6
# so we can populate a segments list
# we can then create the digits : 
# d0 = 3 + 1 + 2 + 6 + 7 + 4
# d1 = 1 + 2
# d2 = 3 + 1 + 5 + 7 + 6
# d3 = 3 + 1 + 2 + 6 + 5
# d4 = 1 + 2 + 4 + 5
# d5 = 3 + 2 + 6 + 4 + 5
# d6 = 2 + 6 + 7 + 4 + 5
# d7 = 3 + 1 + 2
# d8 = 3 + 1 + 2 + 6 + 7 + 4 + 5
# d9 = 3 + 1  +2 + 4 + 5

proc Part02() =
    Initialise
    var myResult : Long
    var mySignal : IterItems = IterItems(s.Signals)
    Do
    
    	var mySignalsVsDigit : Hkvp = GetSignalMap(mySignal.Item)
        Fmt.Dbg "{0},{1}", mySignalsVsDigit.Keys, mySignalsVsDigit.Items
        var myCode : Long = 0
        
        var myI : IterItems = IterItems(s.Digits.Item(mySignal.Index))
        Do
        	DoEvents
            
            var myItem : String = Strs.Sort(myI.Item)
            If mySignalsVsDigit.HoldsKey(myItem) Then
            	
                myCode = myCode * 10 + mySignalsVsDigit.Item(myItem)
                Debug.Print mySignal.Index, myItem, myI.Item, myCode
                Else
                Debug.Print mySignal.Index, myItem, myI.Item
            End If
            
        Loop While myI.MoveNext
        Debug.Print myCode
        myResult += myCode
        
    Loop While mySignal.MoveNext
    
    Fmt.Dbg "The answer to Day {0} part 2 is {1}.  Found is {2}", VBA.Mid$(InputData, 4, 2), "xxx", myResult
    
End Sub

#fdgacbe cefdb cefbgd gcbe: 8394
#fcgedb cgb dgebacf gc: 9781
#cg cg fdcagb cbg: 1197
#efabcd cedba gadfec cb: 9361
#gecf egdcabf bgf bfgea: 4873
#gebdcfa ecba ca fadegcb: 8418
#cefg dcbef fcge gbcadfe: 4548
#ed bcgafe cdgba cbgef: 1625
#gbdfcae bgc cg cgb: 8717
#fgae cfgab fg bagce: 4315

Private Function GetSignalMap(ByRef ipConnections : Seq) : Hkvp

    # create a dictionary of segment wire strings with the string length : the key
    var myI : IterItems = IterItems(ipConnections)
    var myConnections : Hkvp = Hkvp.Deb
    Do
        DoEvents
        Debug.Print myI.Item
    	myConnections.Add VBA.Len(myI.Item), myI.Item
    Loop While myI.MoveNext
    
  
    var mySegments(1 To 7) : String
    var myTmp : Seq
    Set myTmp = Seq.Deb(myConnections.Item(2))
    mySegments(1) = myTmp.Item(1)
    mySegments(2) = myTmp.Item(2)
    
    #Get the segments that are different between 1 and 7
    Set myTmp = Seq.Deb(myConnections.Item(2)).InParamOnly(Seq.Deb(myConnections(3)))
    mySegments(3) = myTmp.Item(1)
    
    # Get the segments that are different between 1 and 4
    Set myTmp = Seq.Deb(myConnections.Item(2)).InParamOnly(Seq.Deb(myConnections(4)))
    mySegments(4) = myTmp(1)
    mySegments(5) = myTmp(2)
    
    #Get the segments that are different between the merge of 7 and 4 , and 8
    Set myTmp = Seq.Deb(myConnections.Item(3)).Merge(Seq.Deb(myConnections(4))).InParamOnly(Seq.Deb(myConnections(7)))
    mySegments(6) = myTmp(1)
    mySegments(7) = myTmp(2)
    
    # now create a map of segments vs number
    var myDigits : Hkvp = Hkvp.Deb
    myDigits.Add GetSegmentCode(mySegments, 3, 1, 2, 6, 7, 4), 0
    myDigits.Add GetSegmentCode(mySegments, 1, 2), 1
    myDigits.Add GetSegmentCode(mySegments, 3, 1, 5, 7, 6), 2
    myDigits.Add GetSegmentCode(mySegments, 3, 1, 2, 6, 5), 3
    myDigits.Add GetSegmentCode(mySegments, 1, 2, 4, 5), 4
    myDigits.Add GetSegmentCode(mySegments, 3, 2, 6, 4, 5), 5
    myDigits.Add GetSegmentCode(mySegments, 2, 6, 7, 4, 5), 6
    myDigits.Add GetSegmentCode(mySegments, 3, 1, 2), 7
    myDigits.Add GetSegmentCode(mySegments, 3, 1, 2, 6, 7, 4, 5), 8
    myDigits.Add GetSegmentCode(mySegments, 3, 1, 2, 4, 5), 9
    
    Return myDigits
    
End Function

Public Function GetSegmentCode(ByRef ipSegments : Variant, ParamArray ipElements() : Variant) : String
	
    var myResult : String
    var myElement : Variant
	For Each myElement In ipElements
    	myResult &= ipSegments(myElement)
	Next
    myResult = Strs.Sort(myResult)
    Return myResult
End Function

proc *Execute() =
        
        Part01
        Part02