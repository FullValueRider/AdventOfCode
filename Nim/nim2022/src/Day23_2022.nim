import strutils
import sequtils
#import strformat
import tables
#import parseutils
import sets
import ../../AoCLib/chars
#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

const AoC2022                                   = "C:\\Users\\slayc\\source\\repos\\AdventOfCode\\RawData"
const Today                                     = "\\2022\\Day23Test1.txt"
const ELF                                       : string = chars.twHash
const NORTH                                     : string = "N"
const SOUTH                                     : string = "S"
const WEST                                      : string = "W"
const EAST                                      : string = "E"

type 
    State = object
        data                                    : Table[Coord,int]
        directions                              : seq[string]
        proposed                                : Table[Coord,seq[Coord]]

    Coord = object 
        X*                                       :int
        Y*                                       :int

 

var s                          : State

iterator reverse*[T](a: seq[T]): T {.inline.} =
    var i = len(a) - 1
    while i > -1:
        yield a[i]
        dec(i)

proc rotL[T]( ipSeq : var seq[T]) =
    ipseq.add(ipSeq[0])
    ipSeq.del(0)
    
# proc merge[T](ipHost, ipInput: var seq[var T]): seq[T] =
#     for myItem in ipInput.items:
#         ipHost.add(myItem)
#     return ipHost
    
proc inBoth[T](ipHost, ipTest: seq[T]) : seq[T] =
    for myCoord in ipHost.items:
        if ipTest.contains(myCoord):
            result.add(myCoord)

proc neighbours(ipCoord : Coord): seq[Coord] =
    var mySeq :seq[Coord] = @[]
    mySeq.add(Coord(X: ipCoord.X - 1, Y:ipCoord.Y + 1 ))
    mySeq.add(Coord(X: ipCoord.X - 1, Y:ipCoord.Y + 1 ))
    mySeq.add(Coord(X: ipCoord.X - 1, Y:ipCoord.Y - 1 ))
    mySeq.add(Coord(X: ipCoord.X, Y:ipCoord.Y + 1 ))
    mySeq.add(Coord(X: ipCoord.X, Y:ipCoord.Y - 1 ))
    mySeq.add(Coord(X: ipCoord.X + 1, Y:ipCoord.Y + 1))
    mySeq.add(Coord(X: ipCoord.X + 1, Y:ipCoord.Y ))
    mySeq.add(Coord(X: ipCoord.X + 1, Y:ipCoord.Y - 1 ))
    return mySeq
    

proc HasNoNeighbours(ipCoord: Coord , ipElves: Table[Coord,int] ): bool =
    
    var myNeighbours = ipCoord.neighbours
    var myBoth = len( ipElves.keys.toSeq.inBoth( myNeighbours))
    return myBoth == 0


proc initialise() =
    
    #[
    We need to remember that info read in will be upside down copared to how it appears
    on the screen.  I.e. the last item in the array is row 0 and the first item in the array is row n
    So we reverse the array before splitting into chars
    ]#
    var myData: seq[string] = (AoC2022 & Today).lines.toSeq.reverse.toSeq

    for myRow in countup( myData.low,myData.high):        
        for myCol in countup(myData[myRow].low, myData[myRow].high):
            
            if $myData[myRow][myCol] == ELF :
                var myKey =  Coord(X:myCol,Y:myRow)
                s.data[myKey] = 0            

        s.directions = @["N", "S", "W", "E"]
     

proc CanMoveNorth(ipCoord: Coord , ipElves: Table[Coord,int] ): bool =

    if ipElves.hasKey(Coord(X:ipCoord.X - 1, Y:ipCoord.Y + 1)) :
         return false

    if ipElves.hasKey(Coord(X: ipCoord.X, Y: ipCoord.Y + 1)) :
        return false
        
    if ipElves.hasKey(Coord(X: ipCoord.X + 1,Y: ipCoord.Y + 1)) :
        return false
    
    return true


proc CanMoveSouth(ipCoord: Coord , ipElves:Table[Coord,int] ): bool =
    
    if ipElves.hasKey(Coord(X: ipCoord.X - 1, Y: ipCoord.Y - 1)) : 
        return false
    if ipElves.hasKey(Coord(X: ipCoord.X, Y: ipCoord.Y - 1)) : 
        return false
    if ipElves.hasKey(Coord(X: ipCoord.X + 1, Y: ipCoord.Y - 1)) : 
        return false
    return true


proc CanMoveEast*(ipCoord:Coord , ipElves: Table[Coord,int] ): bool =
   
    if ipElves.hasKey(Coord(X: ipCoord.X + 1, Y: ipCoord.Y + 1)) :
        return false
    if ipElves.hasKey(Coord(X: ipCoord.X + 1, Y: ipCoord.Y)) :
        return false
    if ipElves.hasKey(Coord(X: ipCoord.X + 1,Y: ipCoord.Y - 1)) : 
        return false
    return true


proc CanMoveWest*(ipCoord: Coord , ipElves: Table[Coord,int] ): bool =

    if ipElves.hasKey(Coord(X: ipCoord.X - 1, Y: ipCoord.Y + 1)) :
        return false
    if ipElves.hasKey(Coord(X: ipCoord.X - 1, Y: ipCoord.Y)) : 
        return false
    if ipElves.hasKey(Coord(X: ipCoord.X - 1, Y: ipCoord.Y - 1)) : 
        return false
    return true


proc GetMoveCoord*(ipCoord: Coord , ipDirections: seq[string] , ipElves: Table[Coord,int]): Coord =
    var myCoord: Coord = Coord(X:0,Y:0)
    for myDirections in ipDirections:
   
        case myDirections  #Note: Nim requires constant expressions for each of
            of NORTH :
                if CanMoveNorth(ipCoord, ipElves) :
                    myCoord.Y += 1
                    break
                
            of SOUTH :
                if CanMoveSouth(ipCoord, ipElves) :
                    myCoord.Y -= 1
                    break
                
            of WEST :
                if CanMoveWest(ipCoord, ipElves) :
                    myCoord.X -= 1
                    break
                
            of EAST :
                if CanMoveEast(ipCoord, ipElves) :
                    myCoord.X += 1
                    break
                
    return myCoord
    


proc GetAreaSize(ipElves : var Table[Coord, int] ): int =

    var myMinX: int = int.high
    var myMinY: int = int.high
    var myMaxX: int = int.low
    var myMaxY: int = int.low
    
    for myCoord, myItem in ipElves.mpairs:

        myMinX = if myCoord.X < myMinX: myCoord.X else: myMinX
        myMinY = if myCoord.Y < myMinY: myCoord.Y else: myMinY
        myMaxX = if myCoord.X > myMaxX: myCoord.X else: myMaxX
        myMaxY = if myCoord.Y > myMaxY: myCoord.Y else: myMaxY

    return ((myMaxX - myMinX) + 1) * ((myMaxY - myMinY) + 1)
    
proc part01() =

    initialise()
    var myNewElves: Table[Coord,int] = s.data
    var myOldElves: Table[Coord, int]
    #var myCount: int
    for myCount in [1..10]:
        myOldElves = myNewElves
        myNewElves.clear

        for myCoord, myElves in myOldElves.mpairs:
            s.proposed.clear
       
            var myOldCoord = myCoord
            var myNewCoord:Coord
            if HasNoNeighbours(myOldCoord, myOldElves) :
                myNewCoord = myOldCoord
            else:
                myNewCoord = GetMoveCoord(myOldCoord, s.directions, myOldElves)
            
            
            if not s.proposed.hasKey(myNewCoord) :
                s.proposed[myNewCoord] = newSeq[Coord]()
            
            s.proposed[myNewCoord].add( myOldCoord)
            
        # now do the moves
        for myNewCoord, myCoords in s.proposed.mpairs:
           
            if myCoords.len == 1 :
                myNewElves[myNewCoord] = 0
                    
            else:
                for myItem in myCoords.items:
                    myNewElves[myItem] = 0
            
        s.directions.rotL

    
    var myResult: int = GetAreaSize(myNewElves) - myNewElves.len
    
    #echo fmt Fmt.Text("The answer to Day {0} part 1 is {1}.  Found is {2}", Mid$(Today, 10, 2), "XXXXXX", myResult)
    var myOutput = "The answer to Day" & $Today[11..12] & "Part 01 is " & "xxxxxxxx" &  " Found is " & $myResult   
    echo myOutput



proc part02() =

    initialise()
    
    
    var myResult = 0
            
    #echo fmt "The answer to Day {0} part 2 is {1}.  Found is {2}", Mid$(Today, 10, 2), "YYYYYY", myResult
    #echo fmt "The answer to Day {Today[11..12} part 1 is 145167969204648.  Found is {myresult}"
    var myOutput = "The answer to Day" & $Today[11..12] & "Part 02 is " & "yyyyyyy" &  " Found is " & $myResult   
    echo myOutput

proc execute*() =
    
    part01()
    #Part02
        

