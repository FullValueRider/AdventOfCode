import tables
import strformat
import sequtils
import strutils
import tools
import ../../AoCLib/src/constants 
#========1=========2=========3=========4=========5=========6=========7=========8========9=========A=========B=========C


const Today                                     : string = "\\2022\\Day03.txt"


type 
    State = object 
        Data                                    : seq[seq[char]]
        Priority                                : Table[char,int]


var
    s                          : State


proc initialise() =

    var myData =
        (AoCData & Today)
            .lines
            .toSeq
            .mapIt(it.items.toSeq)
    s = State( Data: myData )

    var myCounter = 1
    for myChar in 'a' .. 'z':
        s.Priority[myChar]=myCounter
        myCounter += 1
    for myChar in 'A' .. 'Z':
        s.Priority[myChar]=myCounter
        myCounter += 1


proc part01() =

    initialise()
    var myPriorities: int = 0
    for myRucksack in s.Data:
        var myHalves = myrucksack.splitat(myRucksack.len div 2)
        myPriorities += s.Priority[myHalves[0].inBoth(myHalves[1]).first]
   
    var myResult: int = myPriorities

    echo fmt"The answer to Day {Today[9..11]} Part 01 is 7826.  Found is {myResult}"


proc part02() =

    initialise()
    var myPriorities: int = 0
    for myRucksack in countup(0,s.Data.high,3):
        myPriorities += s.Priority[s.Data[myRucksack].inBoth( s.Data[myRucksack+1]).inBoth( s.Data[myRucksack+2]).first]
    
    var myResult: int = myPriorities
            
    echo fmt"The answer to Day {Today[9..11]} Part 02 is 2577.  Found is {myResult}"

proc execute*() = 
    part01()
    part02()


