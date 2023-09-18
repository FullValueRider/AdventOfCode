import strutils
import sequtils
import strformat
import math
#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
const Today = "\\2022\\Day18.txt"
const AoC2022 = "C:\\Users\\slayc\\source\\repos\\AdventOfCode\\RawData"

type
    State = object

        Data                            : seq[string]
        Spaces                          : seq[string]
        Minx                            : int
        MaxX                            : int
        MinY                            : int
        MaxY                            : int
        MinZ                            : int
        MaxZ                            : int
    

var s                                   : State

#[
The gotcha for day 18 is that voids may be bigger than 1 cube
so searching for spaces that are surrounded on all sides will not
work for spaces inside a void.
Therefore we need to workfrom the outside to calculate a number of #not cubes#  not in the droplet then
the size of the voids inside the droplet is 
    Total area - non droplet area - droplet area
]#

proc ToCoordVar(ipcoord: string ): (int,int,int) =
    var mycoord = ipcoord.split(",")
    return (parseInt(mycoord[0]), parseInt(mycoord[1]), parseInt(mycoord[2]))        

proc SetUpBoundariesOfCubeEnclosingDroplet*() =
    
    s.Minx = 2 ^ 15
    s.MaxX = -2 ^ 15
    s.MinY = 2 ^ 15
    s.MaxY = -2 ^ 15
    s.MinZ = 2 ^ 15
    s.MaxZ = -2 ^ 15
    
    for myDroplets in s.Data:
        var mydroplet  = ToCoordVar(mydroplets)
        s.Minx = min(s.Minx, mydroplet[0])
        s.MaxX = max(s.MaxX, mydroplet[0])
        s.MinY = min(s.Minx, mydroplet[1])
        s.MaxY = max(s.MaxX, mydroplet[1])
        s.MinZ = min(s.Minx, mydroplet[2])
        s.MaxZ = max(s.MaxX, mydroplet[2])
        
    #make the surrounding cube bigger than the bounds of the droplet so
    #we don#t run intoissues of blocked volumes due to the shape of the droplet
    s.Minx -= 2
    s.MaxX += 2
    s.MinY -= 2
    s.MaxY += 2
    s.MinZ -= 2
    s.MaxZ += 2

    


proc Initialise() =

    s.Data  = (AoC2022 & Today).lines.toSeq
    s.Spaces = @[]
    
    
proc ToCoordStr(ipX: int, ipY: int, ipZ: int): string =
    return fmt"{ipX},{ipY},{ipZ}"


proc GetNeighbours(ipCoord: string ): seq[string] =

    var myCoord = ToCoordVar(ipCoord)
    
    var myNeighbours: seq[string] = @[]
    myNeighbours.add( ToCoordStr(myCoord[0] + 1, myCoord[1], myCoord[2]))
    myNeighbours.add( ToCoordStr(myCoord[0] - 1, myCoord[1], myCoord[2]))
    myNeighbours.add( ToCoordStr(myCoord[0], myCoord[1] + 1, myCoord[2]))
    myNeighbours.add( ToCoordStr(myCoord[0], myCoord[1] - 1, myCoord[2]))
    myNeighbours.add( ToCoordStr(myCoord[0], myCoord[1], myCoord[2] + 1))
    myNeighbours.add( ToCoordStr(myCoord[0], myCoord[1], myCoord[2] - 1))
        
    return myNeighbours
    

proc TestCube(ipCoord: string ): int =
    
    if s.Data.contains(ipCoord) :
        return 0
    if s.Spaces.contains(ipCoord) == false:
        s.Spaces.add( ipCoord)
    return 1


proc CountMissingNeighbours(ipCoord: string): int =
    
    var myCount: int = 0
    for  myNeighbours in GetNeighbours(ipCoord):
        myCount += TestCube(myNeighbours)
    return myCount
    


proc InBounds(ipX: int, ipY: int, ipZ: int): bool =
    if ipX >= s.Minx and ipX <= s.MaxX :
        if ipY >= s.MinY and ipY <= s.MaxY :
            if ipX >= s.Minx and ipZ <= s.MaxZ :
                return true
    return false

proc GetBoundedNeighbours(ipCoord:  string ): seq[string] =
    var myNeighbours :seq[string]
    var myCoord : (int,int,int) = ToCoordVar(ipCoord)
    
    if InBounds(myCoord[0] + 1, myCoord[1], myCoord[2]) : 
        myNeighbours.add( ToCoordStr(myCoord[0] + 1, myCoord[1], myCoord[2]))
    if InBounds(myCoord[0] - 1, myCoord[1], myCoord[2]) : 
        myNeighbours.add( ToCoordStr(myCoord[0] - 1, myCoord[1], myCoord[2]))
    if InBounds(myCoord[0], myCoord[1] + 1, myCoord[2]) : 
        myNeighbours.add( ToCoordStr(myCoord[0], myCoord[1] + 1, myCoord[2]))
    if InBounds(myCoord[0], myCoord[1] - 1, myCoord[2]) : 
        myNeighbours.add( ToCoordStr(myCoord[0], myCoord[1] - 1, myCoord[2]))
    if InBounds(myCoord[0], myCoord[1], myCoord[2] + 1) :
        myNeighbours.add( ToCoordStr(myCoord[0], myCoord[1], myCoord[2] + 1))
    if InBounds(myCoord[0], myCoord[1], myCoord[2] - 1) : 
        myNeighbours.add( ToCoordStr(myCoord[0], myCoord[1], myCoord[2] - 1))
    return myNeighbours
    
        
proc Part01*() =

    Initialise()
    var mycount: int
    for myCubes in s.Data:
        mycount += CountMissingNeighbours(myCubes)
    
    var myResult: int = mycount
    
    #echo fmt"The answer to Day {Today[11.12]} part 1 is {"XXXXXX"}.  Found is {myResult}"
    var myOutput = "The answer to Day" & $Today[11..12] & "Part 01 is " & "xxxxx" &  " Found is " & $myResult   
    echo myOutput


proc Part02() =

    Initialise()
    
    var myCount: int
    for myCubes in s.Data:
        myCount += CountMissingNeighbours(myCubes)
   
    #echo fmt myCount #correct at this point
            
    SetUpBoundariesOfCubeEnclosingDroplet()
    echo $(s.MaxX - s.Minx + 1) ,$(s.MaxY - s.MinY + 1), $(s.MaxZ - s.MinZ + 1)
    var myTotalVolume: int = (s.MaxX - s.Minx + 1) * (s.MaxY - s.MinY + 1) * (s.MaxZ - s.MinZ + 1)
    echo $myTotalVolume
    var myVoids: seq[string]
    # start the search using a point known to be outside the droplet
    #myVoids.Add ToCoordStr(s.MaxX, s.MaxY, s.MaxZ), 0
    var myQ: seq[string]
    myQ.add(ToCoordStr(s.MaxX, s.MaxY, s.MaxZ)) # future points to investigate
    var myVisited: seq[string]    
    
    while myQ.len > 0:
        echo $myQ.len, $myVoids.len
        var mycoord: string = myQ[0]
        myQ.delete(0)
        if myVisited.contains(mycoord) :
            continue 
        
        myVisited.add( mycoord)
        
        if s.Data.contains(mycoord) :
            continue
        
        if myVoids.contains(mycoord) :
            continue 
        
        myVoids.add(mycoord)
        
        for myNeighbours in GetBoundedNeighbours(mycoord):
            var myneighbour: string = myNeighbours
            if myVisited.contains(myNeighbour) == false:
                myQ.add(myneighbour)
    
    var myResult: int = myTotalVolume - myVoids.len - s.Data.len
            
    var myOutput = "The answer to Day" & $Today[11..12] & "Part 01 is " & "xxxxx" &  " Found is " & $myResult   
    echo myOutput

     
proc execute*() =
    
    Part01()
    Part02()