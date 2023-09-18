import strformat
import ../../AoCLib/src/constants
import checksums/md5

const 
    Today                       = "\\Day04.txt"
    Year                        = "\\2015"



type
  State = object
    data: string
   

var s = new State

proc getFirstPositiveNumberWithMD5StartingWith*( myFirstChars : string) : int =

    var myNum : int = -1
    var myLen : int = myFirstChars.len-1
    
    while true:

        myNum += 1
        
        var myHash : string = getmd5(s.data & $myNum)
        
        if myHash[0..myLen] == myFirstChars:
            break;
        
    return myNum


proc initialise() =
    s.data = readFile(AocData & Year & Today)
    

proc part01() =
    
    initialise()
    var myResult: int = getFirstPositiveNumberWithMD5StartingWith("00000")
    
    echo fmt"The answer to Day {Today[5..6]} part 01 is 117946 .  Found is {myResult}"


proc part02() =
   
    initialise()
    var myResult: int = getFirstPositiveNumberWithMD5StartingWith("000000")
   
    echo fmt"The answer to Day {Today[5..6]} part 2 is 3938038.  Found is {myResult}"        


proc execute*()=
    part01()
    part02()