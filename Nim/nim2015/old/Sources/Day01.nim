import strutils
import strformat
import sequtils
import AoCLib

const Today                     :string = "\\Day01.txt"
const Year                      :string = "\\2015"


const UP = ')'
const DOWN = '('

# The state object is a hangover from VBA where it is seen as good practise
# to encapsulate Module level variables in a User Defined Type.

type
  State = object
    Integers: seq[int]


var s = new State
#s.Integers= toseq(myPath.lines).map(parseInt) 

proc initialise() =
    s.integers = toseq((AocData + Year + Today).lines).map(parseInt) 


proc Part01() =
    
    initialise
    var myResult:int = s.integers.countIt( It = UP ) -s.integers.countit(It = DOWN)

    echo fmt"The answer to Day 01 part 1 is 1711 .  Found is {myResult}"


proc Part02() =
   
    initialise
    var myFloor :int = 0
    var myCount :int = 0
    block myMove:
    for myMove in s.Integers:
        myCount+=1
        myFloor += myMove

        if myfloor < 0 :
            var myResult :int = myCount
            break myMove

    echo fmt"The answer to Day 01 part 2 is 1743.  Found is {myResult}"        


proc Execute*()=
    Part01()
    Part02()