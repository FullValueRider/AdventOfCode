import strformat
import sequtils
import strutils
import AoCLib
import math
import system\iterators
import chars
import sequtils
import strutils
import std/tables
import std/hashes

const InputData :                    string = "Day14.txt"

   
type
    State = object
        Polymer:                        string
        Pairs:                          Table[string,string]

var s = new  State

proc Initialise() =

    var myStrings: seq[string] =split(readfile(RawDataPath2021 & InputData),"\r\n\r\n")
    s.Polymer = myStrings[0]

    var myPairs : seq[seq[string]] = 
        split(myStrings[0], "\r\n")
            .mapIt(it.multireplace((chars.twSpace, chars.twNoString))) 
            .mapIt(it.split("->"))
                
    for myPair in myPairs:
        s.Pairs[myPair[0]]=mypair[1]
    
    
    
proc Part01() =

    Initialise()
    
    # First generate the polymer after x steps
    var myPolymer:string = s.Polymer
    
    
    for mysteps in 1..10:
        var myNew: seq[char]= @[]
        for myIndex, myBase in myPolymer.pairs:
        
            var myFirstChar :char = myBase
            if myIndex<myPolymer.high:
                var mySecondChar :char = mypolymer[myIndex + 1]
                var myKey :string = $myFirstChar & $mySecondChar
                var myInsertChar :string = s.Pairs[myKey]
                
                myNew.add myFirstChar
                myNew.add myInsertChar
            else:
                myNew.add myFirstChar
        
        myPolymer = $myNew
        
    # now compile a histogram of the items in the polymer

    var myHist=toCountTable(myPolymer)
    # var myHist = newTable[char,int]()
    # for myChar in myPolymer:
    #     if myHist.hasKey(myChar):
    #         myHist[mychar]+=1
    #     else:
    #         myHist[mychar]=1

        var mymax: seq[int] = myHist.
    var myResult:int  = @[].foldl(a>b) - myHist.mvalues.foldr(a<b)
    echo fmt"The answer..Day {InputData[ 4..5]} part 1 is {"xxxx"}.  Found is {myResult}"
    

    
proc Part02() =

        Initialise
        
        # First generate the polymer after x steps
        var myPolymer As Seq = s.Polymer.Clone
        
        var mysteps As Long
        for mysteps = 1..40
            var myBase As IterItems = IterItems(myPolymer)
            var myNew As Seq = Seq.Deb
            Do
            
                var myFirstChar As String = myBase.Item
                If myBase.HasNext :
                    var mySecondChar As String = myBase.Item(1)
                    var myKey As String = myFirstChar & mySecondChar
                    var myInsertChar As String = s.Pairs.Item(myKey)
                    
                    myNew.Add myFirstChar
                    myNew.Add myInsertChar
                Else
                	myNew.Add myFirstChar
                End If
            
            Loop While myBase.MoveNext
            
            Set myPolymer = myNew
            
        Next
        
        # now compile a histogram of the items in the polumer
        Set myBase = IterItems(myPolymer)
        var myHist As Hkvp = Hkvp.Deb
        Do
        	If myHist.LacksKey(myBase.Item) :
                myHist.Add myBase.Item, 1
            End If
            myHist.Item(myBase.Item) += 1
            
        Loop While myBase.MoveNext
    
        var myResult As Long = myHist.ReduceIt(rdMax.Deb) - myHist.ReduceIt(rdMin.Deb)
     
        
        
        Fmt.Dbg "The answer..Day {0} part 2 is {1}.  Found is output below", VBA.Mid$(InputData, 4, 2), "xxxx"
        
    End Sub

proc Execute() =
    Part01()
    Part02()