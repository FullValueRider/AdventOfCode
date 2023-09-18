import strutils
import sequtils
import tables
import ../../AoClib/src/chars
    
const MY_TYPENAME: string = "Coord"
    
type 
    XYMovement = object
        xAxis: int
        yAxis: int

    Bounds = object  
        minX : int
        minY :int                                
        maxX :int                                  
        maxY :int 

    MapType = enum  
        OriginTopLeft
        OriginBottomLeft

    CompassPoints = enum
        Four = 4
        Eight = 8

    # nim requires that enums assigned a value are listed in ascending order
    Direction = enum 

        North = 0
        
        Forward = 1
        Backward = 2
        Left = 3
        Right = 4

        
        NE = 45
        East = 90
        SE = 135
        South = 180
        SW = 225
        West = 270
        NW = 315
        
    Properties = object 
        currentX                : int 
        currentY                : int                               
        
        direction               : Direction                            
        compassPoints           : CompassPoints                          
        
        firstRepeatVisit        : string                      
        mostRecentRepeatVisit   : string                 
        
        repeatVisit             : bool                         
        heading                 : string                              
        mapType                 : MapType  

    State = object   
        originX                 : int                               
        originY                 : int                              
        
        directionMap            : Table[string,Direction]                          
        movementMap             : Table[int, array[0..1,int]]                          
        
        mover                   : XYMovement
        turnAngle               : int
        
        visited                 : Table[string,int]
        track                   : seq[string]
        
        isBounded               : bool
        limits                  : Bounds

    Coord* = object
        s: State
        p: Properties


proc populatemovementMap(me : var Coord) =
    
    if me.p.mapType == MapType.OriginBottomLeft:
        
        me.s.movementMap[Direction.North.ord]  = [0, 1]
        me.s.movementMap[Direction.NE.ord]     = [1, 1]
        me.s.movementMap[Direction.East.ord]   = [1, 0]
        me.s.movementMap[Direction.SE.ord]     = [1, -1]
        me.s.movementMap[Direction.South.ord]  = [0, -1]
        me.s.movementMap[Direction.SW.ord]     = [-1, -1]
        me.s.movementMap[Direction.West.ord]   = [-1, 0]
        me.s.movementMap[Direction.NW.ord]     = [-1, 1]
        
    else:
        
        me.s.movementMap[Direction.North.ord]  = [0, -1]
        me.s.movementMap[Direction.NE.ord]     = [1, -1]
        me.s.movementMap[Direction.East.ord]   = [1, 0]
        me.s.movementMap[Direction.SE.ord]     = [1, 1]
        me.s.movementMap[Direction.South.ord]  = [0, 1]
        me.s.movementMap[Direction.SW.ord]     = [-1, 1]
        me.s.movementMap[Direction.West.ord]   = [-1, 0]
        me.s.movementMap[Direction.NW.ord]     = [-1, -1]
            
   

proc constructInstance( ipX :int, ipY :int, ipCompassPoints : CompassPoints, ipMapType : MapType ): Coord =

    # if ipCompassPoints != Four or ipCompassPoints != Eight:
    #     raise newException(OSError, "twLib.Coord.ConstructInstance\n\nCompass points must be 4 or 8")
    
    result = Coord( s:
        State(

            originX                 : ipX,                               
            originY                 : ipY, 

            
            
            directionMap            : {
                "n": Direction.North,
                "north": Direction.North,
                "u": Direction.North,
                "up": Direction.North,
                "^": Direction.North,
                
                "s": Direction.South,
                "south": Direction.South,
                "d": Direction.South,
                "down": Direction.South,
                "v": Direction.South,
                
                "w": Direction.West,
                "west": Direction.West,
                "<": Direction.West,
                
                "e": Direction.East,
                "east": Direction.East,
                ">": Direction.East,
            
            
                "nw": Direction.NW,
                "ne": Direction.NE,
                "se": Direction.SE,
                "sw": Direction.SW,
            
                "f": Direction.Forward,
                "forward": Direction.Forward,
                "forwards": Direction.Forward,
            
                "b": Direction.Backward,
                "back": Direction.Backward,
                "backward": Direction.Backward,
                "backwards": Direction.Backward,
            
                "l": Direction.Left,
                "left": Direction.Left,
            
                "r": Direction.Right,
                "right": Direction.Right}.toTable,    

            movementMap             : initTable[int, array[0..1,int]](),                       
            
            mover                   : XYMovement(xAxis : 0, yAxis : 0),
            turnAngle               : 90,
            
            visited                 : {"0,0":1}.toTable,
            track                   : @["0,0"],
            
            isBounded               : false,
            limits                  : Bounds(minX : 0, minY :0, maxX :0, maxY :0 ),

        ),

        p: Properties(

            currentX                : ipX,
            currentY                : ipY,                        
            
            direction               : Direction.North,                         
            compassPoints           : ipCompassPoints,                          
            
            firstRepeatVisit        : "",                   
            mostRecentRepeatVisit   : "",                
                                
            
            repeatVisit             : false,
            heading                 : "",                              
            mapType                 : ipMapType, 
        )
    )
    
    result.p.compassPoints= if ipCompassPoints == CompassPoints.Four: CompassPoints.Four else: COmpassPoints.Eight
   
    populatemovementMap(result)
    
     #Direction moved to go north depends on Maptype
    result.s.mover.xAxis = result.s.movementMap[North.ord][0]
    result.s.mover.yAxis = result.s.movementMap[North.ord][1]
    
    
    result.s.turnAngle = if result.p.compassPoints == Four: 90 else: 45
        
    result.s.isBounded = false
    
    return result

# proc newCoord*(): Coord =
#     return constructInstance(0, 0, Four, MapType.OriginBottomLeft)
    

proc newCoord*( ipX :int=0, ipY : int=0,  ipCompassPoints : CompassPoints = CompassPoints.Four, ipMapType : MapType = OriginBottomLeft): Coord =
    return constructInstance(ipX, ipY, ipCompassPoints, ipMapType)
    

proc newCoord*( ipXYCoord : string,  ipCompassPoints : CompassPoints = CompassPoints.Four, ipMapType : MapType = OriginBottomLeft): Coord =
    var myCoord : seq[int] = ipXYCoord.split( chars.twComma).mapIt(it.parseInt)
    return constructInstance(myCoord[0], myCoord[1], ipCompassPoints, ipMapType)


proc turnLeft*(me: var  Coord): Coord =
    
    me.p.direction -= me.s.turnAngle
    if me.p.direction < Direction.North:
        me.p.direction = if me.p.compassPoints == CompassPoints.Four: Direction.West else: Direction.NW

    var myMovement  = me.s.movementMap[me.p.direction]
    me.s.mover.xAxis = myMovement[0]
    me.s.mover.yAxis = myMovement[1]
    return me

    
proc turnRight*( me: var Coord): Coord =

    me.p.direction += me.s.turnAngle
    if me.p.direction >= 360:
        me.p.direction =  North.ord 
    
    var myMovement  = me.s.movementMap[me.p.direction]
    me.s.mover.xAxis = myMovement[0]
    me.s.mover.yAxis = myMovement[1]
    return me



proc coordOfFirstRepeatVisit*(me: var Coord) : string =
    return me.p.firstRepeatVisit


proc coordOfMostRecentRepeatVisit*(me : var Coord) : string =
    return me.p.mostRecentRepeatVisit

proc isAtOrigin*(me : var Coord) : bool =
    return (me.s.originX == me.p.currentX) and (me.s.originY == me.p.currentY)


proc reset*(me: var Coord) : Coord =
    me.p.currentX = me.s.originX
    me.p.currentY = me.s.originY
    me.s.track.delete(1..<1)
    me.s.visited.clear
    return me


proc x*(me :  Coord): int =
    return me.p.currentX


proc y*(me :  Coord): int =
    return me.p.currentY


proc toString*(me : Coord) : string =
    return $me.p.currentX & chars.twComma & $me.p.currentY


proc enforceBounds*(me : var Coord) : Coord =

    if me.p.currentX > me.s.limits.maxX : me.p.currentX = me.s.limits.maxX
    if me.p.currentX < me.s.limits.minX : me.p.currentX = me.s.limits.minX
    if me.p.currentY > me.s.limits.maxY : me.p.currentY = me.s.limits.maxY
    if me.p.currentY < me.s.limits.minY : me.p.currentY = me.s.limits.minY
    
    return me


proc move*(me: var Coord, ipDirection : string, ipDistance :int = 1) =

   # echo "move: ipDirection is", $ipDirection
    
    var myDirection:Direction = me.s.directionMap[ipDirection.tolower]
    
    case myDirection
    
        of North, NE, East, SE, South, SW, West, NW:
        
            me.s.mover.xAxis = me.s.movementMap[myDirection.ord][0]
            me.s.mover.yAxis = me.s.movementMap[myDirection.ord][1]
            
        of Direction.Left:
        
            me=turnLeft(me)
            
            
        of Direction.Right:
        
            me=turnRight(me)
            
        of Direction.Forward:
        
            discard
            
        of Direction.Backward:
        
            me.s.mover.xAxis = -me.s.mover.xAxis
            me.s.mover.yAxis = -me.s.mover.yAxis
            
    
    
    if ipDistance < 0 :
        me.s.mover.xAxis = -me.s.mover.xAxis
        me.s.mover.yAxis = -me.s.mover.yAxis
    
    
    
    for mySteps in [1 .. ipDistance.abs]:
    
        me.p.currentX += me.s.mover.xAxis
        me.p.currentY += me.s.mover.yAxis
            
        if me.s.isBounded == true:
            me = enforceBounds(me)
       
     
        
        var mylocation : string = $me.p.currentX & chars.twComma & $me.p.currentY
        
        if me.s.track[me.s.track.high] == mylocation:
            continue
       
        
        me.s.track.add mylocation
        
        if me.s.visited.hasKey(mylocation):
            me.s.visited[mylocation] += 1
            
            if me.p.firstRepeatVisit.len == 0:
                me.p.firstRepeatVisit = mylocation
            
            me.p.mostRecentRepeatVisit = mylocation
            
        else:
        
            me.s.visited[mylocation]= 1
            




proc setBounds*(me: var Coord, ipminX : int,  ipminY : int,  ipmaxX : int,  ipmaxY : int) : Coord =

        me.s.isBounded = true
        
        me.s.limits.minX = ipminX
        me.s.limits.minY = ipminY
        me.s.limits.maxX = ipmaxX
        me.s.limits.maxY = ipmaxY
    
        return me



proc isBounded*(me : Coord) : bool =
    return me.s.isBounded



proc manhatten*(me : Coord) : int =
    return abs(me.p.currentX - me.s.originX) + abs(me.p.currentY - me.s.originY)

proc manhatten(me: Coord, ipX : int, ipY : int) : int =
    return abs(ipX - me.s.originX) + abs(ipY - me.s.originY)



proc manhatten*(me : Coord, ipCoord :string) : int =
    var mycoord : seq[int] = ipCoord.split( chars.twComma).mapit(it.parseint)
    return manhatten(me, mycoord[0],mycoord[1])



proc visited*(me: Coord) : Table[string,int] =
    return me.s.visited

proc track*(me : Coord) : seq[string] =
    return me.s.track
 



proc compassPoints*(me : Coord): CompassPoints =
    me.p.compassPoints


proc typeName*(me : Coord) : string =
    return me.typeName
