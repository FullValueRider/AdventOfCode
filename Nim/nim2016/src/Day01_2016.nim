import strutils

import strformat
import ../../AoCLib/Coord
import ../../AoCLib/constants
import ../../AoCLib/chars

const 
    Today                       = "\\Day01.txt"
    Year                        = "\\2016"
    

type
  State = object
    data: seq[string]
    

var s = new State


proc initialise() =
    s.data = readFile(AocData & Year & Today).replace(twSpace,twNoString).split(twComma)
    
    
proc part01() =
    
    initialise()
    var myWalker : Coord = newCoord()
    for myMove in s.data:
    
        var mydirection: string = $myMove[0]
        var myDistance: string = $myMove[1..^1]
        myWalker.move( mydirection, myDistance.parseInt)
        
    var myResult: int = myWalker.manhatten
    
    echo fmt"The answer to Day {Today[5..6]} part 01 is 74 .  Found is {myResult}"


proc part02() =
   
    initialise()
    var myResult: int = 0
    
   
    echo fmt"The answer to Day {Today[5..6]} part 2 is 1795.  Found is {myResult}"        


proc execute*()=
    part01()
    part02()