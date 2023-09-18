import tables
import strformat
import sequtils
import strutils
import ../../AoCLib/src/constants 
import ../../AoCLib/src/chars

#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
const Today                                     : string = "\\2022\\Day02.txt"

type 

    State = object 
    Data                                    : seq[seq[string]]
        DeCrypterV1                             : Table[string,string]
        DecrypterV2                             : Table[string,string]
        Scorer                                  : Table[string,int]

var 
    s                          : State

proc initialise() =

    var myData: seq[seq[string]] = 
        (AoCData & Today)
            .lines
            .toSeq
            .mapIt(it.split(" "))

    s = 
        State(
            Data : myData, 
            DeCrypterV1 : {"A": "R", "B": "P", "C": "S", "X": "R", "Y": "P", "Z": "S"}.toTable,
            DecrypterV2 : {"R,X":"S", "R,Y":"R", "R,Z":"P", "P,X":"R", "P,Y": "P", "P,Z":"S", "S,X":"P", "S,Y":"S", "S,Z":"R"}.toTable,
            Scorer : {"R,R":4, "R,P":8, "R,S":3, "P,R":1, "P,P":5, "P,S":9, "S,R":7, "S,P":2, "S,S":6}.toTable,
        )

proc part01() =

    initialise()
    var myResult: int = 
        s.Data
            .mapIt(it.mapIt(s.DeCrypterV1[it][0]))
            .mapIt(it.join(","))
            .mapIt(s.Scorer[it])
            .foldl(a + b)

    echo fmt"The answer to Day {Today[9 .. 11]} Part 01 is 15523.  Found is {myResult}"


proc part02() =

    initialise()

    var myNewGames: seq[string] = @[]

    for myGame in s.Data:
        var myOpponent: string = s.DeCrypterV1[myGame[0]]
        myNewGames.add( myOpponent & chars.twComma & s.DecrypterV2[myOpponent & chars.twComma & myGame[1]])
   
    var myResult: int = myNewGames.mapIt(s.Scorer[it]).foldl(a+b)
            
    echo fmt"The answer to Day {Today[9 .. 11]} part 2 is 15702.  Found is {myResult}"


proc execute*() = 
    part01()
    part02()
            