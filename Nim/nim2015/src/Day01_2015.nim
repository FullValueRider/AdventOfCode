import strutils
import strformat

import ../../AoClib/src/constants


const 
    Today                       = "\\Day01.txt"
    Year                        = "\\2015"
    UP                          = '('
    DOWN                        = ')'

# The state object is a hangover from VBA where it is seen as good practise
# to encapsulate Module level variables in a User Defined Type.

type
  State = object
    data: string
   

var s = new State


proc initialise() =
    s.data = readFile(AocData & Year & Today)


proc part01() =
    
    initialise()
    
    var myResult:int = s.data.count( UP ) - s.data.count( DOWN )
    
    echo fmt"The answer to Day {Today[5..6]} Part 01 is 74 .  Found is {myResult}"


proc part02() =
   
    initialise()
    var myFloor :int = 0
    var myCount :int = 0
    var myResult: int = 0
    
    for myMoves in s.data:
        myCount += 1
        myFloor += (if myMoves == UP: 1 else: -1)

        if myfloor < 0 :
            myResult = myCount
            break
   
    echo fmt"The answer to Day {Today[5..6]} Part 02 is 1795.  Found is {myResult}"        


proc execute*()=
    part01()
    part02()