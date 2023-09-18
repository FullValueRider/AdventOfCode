import ../../AoCLib/src/constants
import strutils
import sequtils
import algorithm
import tools
#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
 

   
const Today                                     : string = "\\2022\\Day01.txt"


type 
    State = object
        Data                                    : seq[seq[int]]

var s                                           : State


proc initialise() =
    var myData: seq[seq[int]] = 
        (AoCData & Today)
            .readFile
            .split("\r\n\r\n")
            .mapIt( it.split("\r\n"))
            .mapIt(it.mapIt(it.parseInt))
    s = State(Data: myData)
    
proc part01() =
    initialise()
    var myResult :int = 
        s.Data
            .mapIt(it.foldl(a + b))
            .sorted
            .last
           
    echo "The answer to Day " & $Today[9 .. 10] & " Part 01 is " & "72017." &  " Found is " & $myResult   
    #echo fmt Fmt.Text("The answer to Day {0} part 1 is {1}.  Found is {2}", Mid$(Today, 10, 2), "72017", myResult)

proc part02() =
    initialise()
    var myResult: int = 
        s.Data
            .mapIt(it.foldl(a + b))
            .sorted[^3 .. ^1]
            .foldl(a + b)
   
    echo "The answer to Day " & $Today[9 .. 10] & " Part 02 is " & "212520." &  " Found is " & $myResult    
    #echo fmt "The answer to Day {0} part 2 is {1}.  Found is {2}", Mid$(Today, 10, 2), "212520", myResult

proc execute*() = 

    part01()
    part02()
    




