import strutils
import sequtils
import strformat

const Today:string = "\\Day10.txt"

proc part01() =

    var myS: string = "1113222113"
    var myNewS : string = ""
    var mycount : int = 1
    var myCurrentChar: string
    while myCount < 41:

        echo "level 1","  ", myCount

        myNewS = ""
        var myCtr :int = 0
        for myChar in myS.items.toSeq.mapIt($it): 
            
            if myCurrentChar.len == 0:
                myCurrentChar = myChar
                myCtr = 1
            elif myChar == myCurrentChar:
                myCtr += 1
            else:
                myNewS = myNewS & $myCtr & myCurrentChar
                myCurrentchar = myChar
                myCtr = 1
                
        myNewS = myNewS & $myCtr & myCurrentChar
       
        myS = myNewS
        myCurrentchar = ""
        myCount += 1
    
    var myResult :int = myNewS.len
    
    echo fmt"The answer to Day {Today[5..6]} part 01 is 252594 .  Found is {myResult}"

proc part02() =

    var myS: string = "1113222113"
    var myNewS : string = ""
    var mycount : int = 1
    var myCurrentChar: string
    while myCount < 51:

        echo "level 1","  ", myCount

        myNewS = ""
        var myCtr :int = 0
        for myChar in myS.items.toSeq.mapIt($it): 
            
            if myCurrentChar.len == 0:
                myCurrentChar = myChar
                myCtr = 1
            elif myChar == myCurrentChar:
                myCtr += 1
            else:
                myNewS = myNewS & $myCtr & myCurrentChar
                myCurrentchar = myChar
                myCtr = 1
                
        myNewS = myNewS & $myCtr & myCurrentChar
       
        myS = myNewS
        myCurrentchar = ""
        myCount += 1
    
    var myResult :int = myNewS.len
    
    echo fmt"The answer to Day {Today[5..6]} part 01 is 3579328 .  Found is {myResult}"


proc execute*() =
        part01()
       ' part02()  

   