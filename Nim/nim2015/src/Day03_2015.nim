import strformat
import sequtils
import tables

import ../../AoCLib/src/constants
import Journey

const 
    Today                       = "\\Day03.txt"
    Year                        = "\\2015"
 
 
type
  State = object
    data: string
    

proc initialise(ipdata :var string) =
    ipData =readFile(AocData & Year & Today)
    

proc part01() =
    var s = State( data: "")
    initialise(s.data)
    var mySanta  = initJourney(s.data)
   
    var myResult: int = mySanta.Visits.len
    echo fmt"The answer to Day {Today[5..6]} part 01 is 2572 .  Found is {myResult}"


proc part02() =
   
    var s = State( data: "")
    initialise(s.data)
    var mySantaMoves: string
    var myRobotMoves: string
   
    for myMoveIndex in countup(s.data.low, s.data.high, 2):
        mySantaMoves &=  $s.data[myMoveIndex]
        myRobotMoves &=  $s.data[myMoveIndex+1]
        
    var mySantaJourney = initJourney(mySantaMoves)
    var myRobotJourney = initjourney(myRobotMoves)
    var myResult  = mySantaJourney.Visits.keys.toSeq.concat(myRobotJourney.Visits.keys.toSeq).deduplicate.len
    echo fmt"The answer to Day {Today[5..6]} part 02 is 2631.  Found is {myResult}"        

 
proc execute*()=
    part01()
    part02()