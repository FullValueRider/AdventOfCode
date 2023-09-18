import sequtils
import tables
import tools
#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

#[
    Ths problem is about constructnbg multiple trails and then selecting the 
    one which is the shortest.

    s.data is a seq of seq where each item represent the height; 
]#

const AoC2022                                   : string = "C:\\Users\\slayc\\source\\repos\\AdventOfCode\\RawData"
const Today                                     : string = "\\2022\\Day12.txt"
const startItem                                 : char = 'S'
const endItem                                   : char = 'E'

type 

    Coord = object 
        col                                     : int
        row                                     : int
        
    State = object 
        data                                    : Table[Coord, char]
        endcoord                                : Coord
        startcoord                              : Coord
        endfound                                : bool
        visited                                 : seq[Coord]


var s :  State 

proc last[T](ipSeq: seq[T]): T =
    return ipSeq[ipSeq.high]

proc initialise() =

    #[
    We need to remember that info read in will be upside down copared to how it appears
    on the screen.  I.e. on screen the last row read in is row 0 on screen 
    and the first row read in is the last row on screen.
    So we reverse the array before splitting into chars
    ]#
    var myData: seq[seq[char]] = (AoC2022 & Today).lines.toseq.reverse.toseq.mapIt(it.items.toseq) 

    s = State(
        data : initTable[Coord,char](),
        endfound: false,
        visited: @[],
    )
  
    # convert myData into Table[Coord,char]
    for myRowIndex, myRow in myData.pairs:
        for myColIndex, myChar in myRow.pairs:
            var myElevation = myChar
            var myCoord = Coord(col: myColIndex, row: myRowIndex)
            
            if myElevation == startItem:
                s.startcoord = myCoord
                myElevation = 'a'

            elif myElevation == endItem:
                s.endcoord = myCoord
                myElevation = 'z'
                
            s.data[ myCoord] = myElevation


proc getAdjacentCoordinates(ipCoord: Coord, ipSearchArea: CompassPoints = Fourway): seq[Coord] =

    let myOffsets : seq[seq[int]] = if ipSearchArea == Fourway: FourWayOffsets else: EightWayOffsets

    for myOffset in myOffsets:
        var myCoord: Coord = Coord(col: ipCoord.col + myoffset[0], row: ipCoord.row + myOffset[1])
        if myCoord in s.data:
           result.add(myCoord)


#@Description("Get the coordinates surrounding the last position if they haven't been visited before 
#in ipTrack and if they are +1 in height or less of the last Coord (coord in ipData")
proc getEligibleSurroundingCoords(ipLastCoord: Coord ): seq[Coord] =
    
    #var myLastCoord: Coord = ipTrack.last
  
    # Get all adjacent points; a sequence of string coordinates using fourway compass points
    # this will only get points at or within the setbounds
    var myAdjacentCoords: seq[Coord] = getAdjacentCoordinates(ipLastCoord, Fourway)
    
    if myAdjacentCoords.len == 0:
        return @[]

    # we now need only those coordinates that are new, i.e. haven't already been visited by ipSTrack
    var myNewCoords: seq[Coord] = myAdjacentCoords.lhsOnly(s.visited)
 
    if myNewCoords.len == 0:
        return @[]
    # now elimiate those new coordinated that are +2 or more higher than than myLastCoord.Item
    # to find coordinates eligible for extending the curren track
    # var myEligibleCoords: seq[Coord] = @[]
    for myNewCoord in myNewCoords:
        if s.data[myNewCoord].int < s.data[ipLastCoord].int + 2:
            s.visited.add(myNewCoord)
            result.add(myNewCoord)
     

proc extendCurrentTrack(ipCurrentTrack: seq[Coord] ): seq[seq[Coord]] =
    
    var myLastCoord : Coord = ipCurrentTrack.last
    var myEligibleAdjacentCoords: seq[Coord] = getEligibleSurroundingCoords(myLastCoord) 
    if myEligibleAdjacentCoords.len == 0:
        return @[]

    for myEligible in myEligibleAdjacentCoords:

        var myExtendedTrack: seq[Coord] = ipCurrentTrack
        myExtendedTrack.add( myEligible )

        result.add( myExtendedTrack )
        # check to see if the most recent updated track ends at the highest levels
        
        if s.data[myLastCoord] == 'z' :
            if myEligible == s.endcoord:
                s.endfound = true
                break
    
    return result


proc part01() =

    initialise()
   
    var myFirstTrack: seq[Coord]  = @[ s.startCoord]
    var myNewTracks: seq[seq[Coord]] = @[myFirstTrack]
    var mycounter: int = 0
    
    while  s.endfound == false and mycounter < 32:  # 32 for debugging 
        
        # echo $mycounter & "    " & $myNewTracks.len
        # mycounter += 1
        var myCurrentTracks: seq[seq[Coord]] = myNewTracks
        myNewTracks = @[]
          
        for myTrack in myCurrentTracks:
            #echo $myTrack
            var mySeq: seq[seq[Coord]] = extendCurrentTrack(myTrack)
           
            myNewTracks &= mySeq 
            
    var myResult: int = myNewTracks.mapIt(it.len).foldl( if a<b: a else: b)-1
    
    var myOutput = "The answer to Day" & $Today[9..10] & " Part 01 is " & "391" &  " Found is " & $myResult
    echo myOutput
    #echo fmt"The answer to Day {0} part 1 is {1}.  Found is {2}", Mid$(Today, 10, 2), "XXXXXX", myResult)
           

proc part02() =
    
    initialise()
    var myStartCoords: Table[Coord,char] = initTable[Coord,char]()
    for myCoord,myChar in s.data:
        if myChar == 'a':
            myStartCoords[myCoord]=myChar

    echo "Starts found = " & $myStartCoords.len   

    var myMin: int = int.high
    var myCounter: int = 0

    for myStartCoord, myChar in myStartCoords.pairs:
        echo $myCounter
        myCounter += 1
        s.endfound = false
        s.visited = @[]
        var myFirstTrack: seq[Coord]  = @[ mystartCoord]
        var myNewTracks: seq[seq[Coord]] = @[myFirstTrack]
        
        while s.endfound == false: 

            if myNewTracks.len == 0:
                break

            var myCurrentTracks: seq[seq[Coord]] = myNewTracks
            myNewTracks = @[]
            
            for myTrack in myCurrentTracks:
                var mySeq: seq[seq[Coord]] = extendCurrentTrack(myTrack)

                #echo "myNewTracks = " & $myNewTracks.len & " mySeq = " & $mySeq.len
                myNewTracks &= mySeq 
                #echo "myNewTracks = " & $myNewTracks.len & " mySeq = " & $mySeq.len

        #echo "myNewTracks = " & $ myNewTracks.len
        if myNewTracks.len > 0:
            var myTrackMin: int = myNewTracks.mapIt(it.len).min  - 1
            if myTrackMin < myMin:
                myMin = myTrackMin

    var myResult: int = myMin
    
    var myOutput = "The answer to Day" & $Today[9..10] & " Part 02 is " & "YYYYYY" &  " Found is " & $myResult
    echo myOutput
    #echo fmt"The answer to Day {0} part 02 is {1}.  Found is {2}", Mid$(Today, 10, 2), "XXXXXX", myResult)

    
proc execute*() = 
    #part01()
    part02()