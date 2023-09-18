import strformat
import sequtils

import ../../AoCLib/src/constants
import Box

const 
    Today             : string = "\\Day02.txt"
    Year              : string = "\\2015"

type
  State = object
    data: seq[string]
# vas s works because we only have one instance of Day02
#var s = new State


proc initialise(): seq[string] =
    return (AocData & Year & Today).lines.toSeq


proc part01() =
    var s = State(data:initialise())
    var myResult : int = 0
    # we don't need to kep the new Box instance, we just need the wrapping size
    for myBox in s.data:
         myresult += newBox(myBox).wrappingSize
    echo fmt"The answer to Day {Today[5..6]} part 01 is 1598415.  Found is {myResult}"


proc part02() =
   
    var s=State(data:initialise())
    var myResult:int = 0
    for myBox in s.data:
        myResult += newBox(myBox).ribbonlength

    echo fmt"The answer to Day {Today[5..6]} part 02 is 3812909.  Found is {myResult}"        


proc execute*()=
    part01()
    part02()