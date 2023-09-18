import ../../AoCLib/src/constants 
import strformat
import tools
#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C


const Today            : string = "\\2022\\Day06.txt"


type 
    State = object 
        Data                                : string



var s                                       : State


proc initialise() =
    s = State(Data : (AoCData & Today).readFile)


proc part01() =

    initialise()

    var myBuffer: seq[char] = @[]
    var myFoundIndex: int

    for myIndex,myChar in s.Data:
        
        if myBuffer.holdsitem(myChar):
            while true:
                if myBuffer.dequeue == myChar:
                    break

        mybuffer.add(mychar)

        if mybuffer.len == 4:
            myfoundindex = myIndex
            break

    var myResult: int = myFoundIndex + 1

    echo fmt"The answer to Day {Today[9 .. 10]} Part 01 is 1356.  Found is {myResult}"


proc part02() =

    initialise()
    # The only change from part 1 was increasing the count from 4 to 14
    var myBuffer: seq[char] = @[]
    var myFoundIndex: int

    for myIndex,myChar in s.Data:
       
        if myBuffer.holdsitem(myChar):
            while true:
                if myBuffer.dequeue == myChar:
                    break

        mybuffer.add(mychar)

        if mybuffer.len == 14:
            myfoundindex = myIndex
            break

                
    var myResult: int = myFoundIndex + 1

    echo fmt"The answer to Day {Today[9 .. 10]} Part 02 is 2564.  Found is {myResult}"

proc execute*() = 
    part01()
    part02()
    

