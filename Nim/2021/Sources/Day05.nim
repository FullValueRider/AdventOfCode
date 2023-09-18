import strformat
import sequtils
import strutils
import tables 
import AoCLib
import VentLinesStartEndCoords 

const InputData = "Day05.txt"

type

    State = object
        Ventlines:seq[VentLineStartEndCoords]
 
   
var s:State

proc Initialise() =
    s=  State( VentLines :
        readfile(RawDataPath2021 & InputData)
            .split("\r\n")
            .mapIt(it.multireplace(("  ","")))
            .mapIt(it.multireplace((" -> ",",")))
            .mapIt( initVentLineStartEndCoords(it))
            )

proc AddVentMapLine( ipVentLine: VentLineStartEndCoords, iopVentMap: var Table[string,int]) =
    
    var mycoords:seq[string] = ipVentLine.GetVentLineCoords
    
    for myCoord in mycoords:
        if iopVentMap.haskey(myCoord):
            iopVentMap[myCoord] = iopVentMap[myCoord] +  1
        
        else:
            iopVentMap[myCoord] = 1


proc BuildVentMap(ipVentLines: seq[VentLineStartEndCoords], ipAllowedLineTypes: seq[VentLineType] ): Table[string,int] =
    var myVentMap = Table[string,int]()
    
    for  myVentLine in ipVentLines:
        if ipAllowedLineTypes.contains(myVentLine.GetLineType()):
            AddVentMapLine( myVentLine, myVentMap)

    return myVentMap



proc Part01() =
    Initialise()

    var myVentLineTypes:seq[VentLineType] = 
        @[
            VentLineType.vlPoint,
            VentLineType.vlHorizontal,
            VentLineType.vlVertical    
        ]
    var myResult: int =BuildVentMap(s.VentLines, myVentLineTypes).values.toseq.countIt( it > 1)
    echo fmt"The answer to Day 05 part 1 is 7085.  Found is {myResult}"


proc Part02() =
    Initialise()
    var myVentLineTypes:seq[VentLineType] = 
        @[
            VentLineType.vlPoint,
            VentLineType.vlHorizontal,
            VentLineType.vlVertical,    
            VentLineType.vlDiagonal
        ]
    var myResult: int =BuildVentMap(s.VentLines, myVentLineTypes).values.toseq.countIt(it > 1)
    echo fmt"The answer to Day 05 part 2 is 20271.  Found is {myResult}"

proc Execute*() =
    Part01()
    Part02()

