import strformat
import sequtils
import strutils
import tables
import ../../AoCLib/src/constants 
import ../../AoCLib/src/chars
import tools
#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

#[
    This solution highlights a significant difference between the behaviour
    of the lookup/Hkvp objects used in VBA and the nim Table.  
    The former, by how they are constructed, have ordered entries, 
    whereas the latter does not.  Thus to get the same 
    behaviour in nim as in VBA we need to use an OrderedTable for s.Boxes.
]#
const Today                                     : string = "\\2022\\Day05.txt"

type
    Instruction = enum
        Count = 0
        MoveFrom = 1
        MoveTo = 2


    State = object 
        Data                                    : seq[seq[string]]
        Boxes                                   : OrderedTable[string, seq[char]]
        Instructions                            : seq[seq[string]]


var 
    s                                           : State

 #@Description("Transpose a seq of seq")
proc transpose[T]( ipSeq : seq[seq[T]]) : seq[seq[T]] =
    
    for myRowIndex in ipSeq.first.low .. ipSeq.first.high:
        result.add( newSeq[T](ipSeq.len))
    
    for myRowIndex,myRow in ipSeq:
        for myColIndex,myCol in myRow:
            result[myColIndex][myRowIndex] = ipSeq[myRowIndex][myColIndex]


proc initialise() =

    var myData: seq[seq[string]] = 
        (AoCData & Today)
            .readFile
            .split("\r\n\r\n")
            .mapit(it.split("\r\n"))

    s = State(  Data : myData,  Instructions: @[], Boxes: initOrderedTable[string,seq[char]]())
    
    var myTransposed  = transpose(s.Data.first.reverse.toSeq.mapit(it.items.toSeq))
    
    
    for myRow in myTransposed:

        var myRowCopy = myRow
        if myRowCopy.first == chars.twSpace:
            continue
        else:
            # remove spaces at the end of the row
            var myKey: string = $myRow.first
            var mySlice = myRow[1 .. ^1]
            while mySlice.last == chars.twSpace:
                discard mySlice.pop
            s.Boxes[myKey] = mySlice

    # now process the instructions for movement
    s.Instructions = 
        s.Data
            .last 
            .mapIt(it.replace("move ", ""))
            .mapIt(it.replace(" from", ""))
            .mapIt(it.replace(" to", ""))
            .mapIt(it.split(chars.twSpace))


proc getLastBoxes(): string =
    for myStack in s.Boxes.values:
        if myStack.len > 0 :
            result &= $myStack.last 


proc part01() =

    initialise()
   
    for myInstruction in s.Instructions:
        
        var myfrom: string = myInstruction[Instruction.MoveFrom.ord]
        var myTo: string = myInstruction[Instruction.MoveTo.ord]
        
        for myMove in 1 .. myInstruction[Instruction.Count.ord].parseInt:
            if s.Boxes[myfrom].len == 0 :
                continue
            var myItem: char = s.Boxes[myfrom].pop
            s.Boxes[myTo].add(myItem)
   
    var myResult: string = getLastBoxes()
    
    echo fmt"The answer to Day {Today[9..10]} Part 01 is FRDSQRRCD.  Found is {myResult}"


proc part02() =

    initialise()
    
    for myinstruction in s.Instructions:

        var myFromStack: seq[char] = s.Boxes[myinstruction[Instruction.MoveFrom.ord]]
        var myToStack: seq[char] = s.Boxes[myinstruction[Instruction.MoveTo.ord]]
        var myCount: int = ($myinstruction[Instruction.Count.ord]).parseInt
        echo fmt"{myFromStack}, {myToStack}, {myCount}"
       
        if myFromStack.len == 0 :
            continue
        
        if myCount >= myFromStack.len :
            
            myToStack  = myToStack.concat(myFromStack)
            myFromStack = @[]

        else:
            
            myToStack = myToStack.concat( myFromStack.splitAt[^(myCount) .. ^1])
        echo fmt"{myFromStack}, {myToStack}"
        echo "\r\n"
        
    var myResult: string = getLastBoxes()

    echo fmt"The answer to Day {Today[9..10]} Part 02 is HRFTQVWNN.  Found is {myResult}"


proc execute*() = 
    part01()
    part02()
    