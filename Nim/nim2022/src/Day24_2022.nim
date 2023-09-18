import strutils
import sequtils
import tables
#import chars
import ../../AoCLib/src/chars
import Tools
#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
const AoC2022                                   = "C:\\Users\\slayc\\source\\repos\\AdventOfCode\\RawData"
const Today                                     = "\\2022\\Day24.txt"
const PATH                                      : string = $twPeriod

type
    Direction = enum
        NORTH = (0, $twHat)
        WEST  = (1, $twLAngle)
        SOUTH = (2, $'v')
        EAST  = (3, $twRAngle)

    Point  = tuple
        col: int
        row: int

    State = object

        Ticker                                  : int

        Start                                   : Point
        Finish                                  : Point
        
        BlizzardMinCol                          : int
        BlizzardMinRow                          : int
        
        BlizzardMaxCol                          : int
        BlizzardMaxRow                          : int
        Gateways                                : seq[Point]

        FutureOffsets                           : seq[Point]
        Blizzards                               : Table[string,seq[Point]]
              
var s : State = State()

proc IsPath(ipPoint :Point): bool =

    if ipPoint in s.Gateways:
        return true
    
    if ipPoint in s.Blizzards[$NORTH]:
        return false
    if ipPoint in s.Blizzards[$SOUTH]:
        return false
    if ipPoint in s.Blizzards[$EAST]:
        return false
    if ipPoint in s.Blizzards[$WEST]:
        return false

    if ipPoint.col < s.BlizzardMinCol:
        return false
    if ipPoint.col > s.BlizzardMaxCol:
        return false
    if ipPoint.row < s.BlizzardMinRow:
        return false
    if ipPoint.row > s.BlizzardMaxRow:
        return false

    return true


proc Initialise() =
  
    s.FutureOffsets = @[(0,0),(0,1),(1,0),(0,-1),(-1,0)]
    var myData :seq[string] =(AoC2022 & Today).lines.toSeq.reverse.toSeq
  
    s.BlizzardMinCol = 1
    s.BlizzardMinRow = 1
  
    s.BlizzardMaxCol = myData.first.high - 1
    s.BlizzardMaxRow = myData.high - 1
    
    s.Blizzards =  {$NORTH: newSeq[Point](),  $SOUTH:  newSeq[Point](), $EAST:  newSeq[Point](), $WEST:  newSeq[Point]() }.toTable  

    for myRow in countdown(s.BlizzardMaxRow, s.BlizzardMinRow):
        for myCol in s.BlizzardMinCol..s.BlizzardMaxCol:

            var myChar = $myData[myRow][mycol]
            case myChar
                of $NORTH, $SOUTH, $EAST, $WEST:
                    s.Blizzards[$myChar].add((myCol,myRow))
                else:
                    discard
    
    s.Finish = (col: myData.first.find(PATH,0,myData.first.high),  row: myData.low)
    s.Start = ( col: myData.last.find(PATH,0,myData.first.high) , row: myData.high)
    s.Gateways = @[]
    s.Gateways.add(s.Start)
    s.Gateways.add(s.Finish)
  
proc HitsBoundary(ipPoint : Point) : bool =
    if ipPoint.col < s.BlizzardMinCol:
        return true
    if ipPoint.col > s.BlizzardMaxCol:
        return true
    if ipPoint.row < s.BlizzardMinRow:
        return true
    if ipPoint.row > s.BlizzardMaxRow:
        return true
    return false


# When checking bounds we have to remember that the entrance and exit
# have the same row coordinates as the the top and bottom walls
# By examining the map we know that there are no blizzards that can enter the start or ending positions
proc MoveBlizzards() =
    
    s.Ticker += 1

    var myS: seq[Point] = @[]
    var myNewPoint : Point

    for myPoint in s.Blizzards[$NORTH].mitems:
        myNewPoint = ( col:myPoint.col, row:myPoint.row + 1)
        if myNewPoint.HitsBoundary:
            myNewPoint.row = s.BlizzardMinRow
        myS.add(myNewPoint)
    s.Blizzards[$NORTH]=myS
    
    myS = @[]
    for myPoint in s.Blizzards[$SOUTH].mitems:
        myNewPoint = ( col:myPoint.col, row:myPoint.row - 1)
        if myNewPoint.HitsBoundary:
            myNewPoint.row = s.BlizzardMaxRow
        myS.add(myNewPoint) 
    s.Blizzards[$SOUTH] = myS

    myS = @[]
    for myPoint in s.Blizzards[$EAST].mitems:
        myNewPoint = ( col:myPoint.col + 1 , row:myPoint.row)
        if myNewPoint.HitsBoundary:
            myNewPoint.col = s.BlizzardMinCol
        myS.add(myNewPoint) 
    s.Blizzards[$EAST] = myS

    myS = @[]
    for myPoint in s.Blizzards[$WEST].mitems:
        myNewPoint = ( col:myPoint.col - 1, row:myPoint.row)
        if myNewPoint.HitsBoundary:
            myNewPoint.col = s.BlizzardMaxCol
        myS.add(myNewPoint) 
    s.Blizzards[$WEST] = myS
    

proc GetNewPathPoints(ipOldPathPoints: seq[Point]): seq[Point] =

    var mySeq : seq[Point] = @[]
    for myMove in ipOldPathPoints:
        for myO in s.FutureOffsets:
            var myPoint :Point = (col: myMove.col + myO[0], row: myMove.row + myO[1])
            if IsPath(myPoint):
                if myPoint notin mySeq:
                    #echo myPoint, "   IsPath"
                    mySeq.add(myPoint) 
    return mySeq

proc Part01() =

    Initialise()
    
    var myPathPoints =  newSeq[Point]()
    myPathPoints.add(s.Start)
   
    while s.Finish notin myPathPoints:
        
        MoveBlizzards()
        echo $s.Ticker & "     " & $myPathPoints.len

        myPathPoints = GetNewPathPoints(myPathPoints)
      
    var myResult: int = s.Ticker 
    
    #echo fmt Fmt.Text("The answer to Day {0} part 1 is {1}.  Found is {2}", Mid$(Today, 10, 2), "XXXXXX", myResult)
    echo "The answer to Day " & $Today[9..10] & " Part 01 is " & "yyyyyyy." &  " Found is " & $myResult   
    


proc Part02*() =

    Initialise()
    
    var myPathPoints =  newSeq[Point]()
    myPathPoints.add(s.Start)
   
    while s.Finish notin myPathPoints:
        
        MoveBlizzards()
        echo $s.Ticker & "     " & $myPathPoints.len

        myPathPoints = GetNewPathPoints(myPathPoints)

    myPathPoints =  newSeq[Point]()
    myPathPoints.add(s.Finish)
   
    while s.Start notin myPathPoints:
        
        MoveBlizzards()
        echo $s.Ticker & "     " & $myPathPoints.len

        myPathPoints = GetNewPathPoints(myPathPoints)

    myPathPoints =  newSeq[Point]()
    myPathPoints.add(s.Start)
   
    while s.Finish notin myPathPoints:
        
        MoveBlizzards()
        echo $s.Ticker & "     " & $myPathPoints.len

        myPathPoints = GetNewPathPoints(myPathPoints)
            
    var myResult: int = s.Ticker 

    #echo fmt "The answer to Day {0} part 2 is {1}.  Found is {2}", Mid$(Today, 10, 2), "YYYYYY", myResult
    echo "The answer to Day" & $Today[9..10] & "Part 02 is " & "yyyyyyy" &  " Found is " & $myResult   


proc Execute*() = 
        
    Part01()
    #Part02()