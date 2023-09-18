import strformat
import sequtils
import strutils
import ../../AoCLib/src/constants 
import ../../AoCLib/src/chars

#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
const Today                                     : string = "\\2022\\Day04.txt"


type 
    State = object 
        Data                                    : seq[seq[string]]


var 
    s                                           : State


proc initialise() =
    s = State( Data : (AoCData & Today).lines.toSeq.mapIt(it.split($chars.twComma)))
    
    

# proc UpgradeToDoubleDigits(ipstring: var string ): string = 

#     var mySeq = ipString.split( chars.twHyphen)

#     if myseq[0].len == 1 :
#         mySeq[0] = "0" & mySeq[0]


#     if mySeq[1].len == 1 :
#         mySeq[1] = "0" & mySeq[1]


#     return Join(mySeq, Char.twHyphen)


proc part01() =

    initialise()

    var myCount: int = 0

    for mySections in s.Data:
    
        var mySection: seq[seq[int]] = mySections.mapit(it.split(chars.twHyphen).mapit(it.parseInt))

        if mySection[0][0] <= mySection[1][0] :
            if mySection[0][1] >= mySection[1][1] :
                myCount += 1
                continue

        if mySection[1][0] <= mySection[0][0] :
            if mySection[1][1] >= mySection[0][1] :
                myCount += 1
                continue

    var myResult: int = myCount

    echo fmt"The answer to Day {Today[9 .. 11]} Part 01 is 657.  Found is {myResult}"




proc part02() =

    initialise()

    var myCount: int = 0

    for mySections in s.Data:
        var mySection: seq[seq[int]] = mySections.mapit(it.split(chars.twHyphen).mapit(it.parseInt))
        
        if mySection[0][0] <= mySection[1][0] :
            if mySection[0][1] >= mySection[1][0] :
                myCount += 1
                continue
            
        if mySection[0][0] <= mySection[1][1] :
            if mySection[0][1] >= mySection[1][1] :
                myCount += 1
                continue

        if mySection[1][0] <= mySection[0][0] :
            if mySection[1][1] >= mySection[0][0] :
                myCount += 1
                continue

        if mySection[1][0] <= mySection[0][1] :
            if mySection[1][1] >= mySection[0][1] :
                myCount += 1
                continue

    var myResult: int = myCount
            
    echo fmt"The answer to Day {Today[9 .. 11]} Part 02 is 938.  Found is {myResult}"

proc execute*() = 
    part01()
    part02()
