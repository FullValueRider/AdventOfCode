import tables
import strutils
import strformat

type
    State = object
        x: int
        y: int
        directions: Table[string, seq[int]]

    Properties = object
        path: seq[string]
        visits : Table[string, int]

    Journey = object 
        s*: State
        p*: Properties


const NORTH                 : string = "^"
const SOUTH                 : string = "v"
const WEST                  : string = "<"
const EAST                  : string = ">"

proc UpdatePath( me: var seq[string], ipCoords: string) =
    #echo fmt"Update path {me.len}"
    me.add(ipCoords)

proc UpdateVisits(me: var Table[string,int], ipCoords :string) = 
    if me.hasKey(ipCoords):
        me[ipCoords]  = me[ipCoords] + 1
    else:
        me[ipCoords] = 1
        


proc initJourney*(ipMoves: string) : Journey =
    var mJ: Journey = Journey(s:State(x:0, y:0, directions: { NORTH: @[0, 1], SOUTH: @[0, -1], EAST: @[-1, 0], WEST: @[1, 0] }.toTable), p:Properties( visits: initTable[string,int](), path: @[]))
  
    for  myMove in ipMoves:
        var myArray = mJ.s.directions[$myMove]
        mJ.s.x +=  myArray[0]  # The ! after the variable tells the compiler this value will never be undefined.
        mJ.s.y +=  myArray[1]
        
        var myCoords: string = fmt"{mJ.s.x},{mJ.s.y}"
        UpdatePath( mJ.p.path, myCoords)
        UpdateVisits( mJ.p.visits, myCoords)

    return mJ


proc Visits*(me : Journey): Table[string,int] =
    return me.p.visits

proc Path*(me: Journey): seq[string] =
    return  me.p.path
    
