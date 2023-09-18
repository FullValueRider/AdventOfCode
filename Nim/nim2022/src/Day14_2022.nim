import strutils
import sequtils
import strformat
import tables
import Tools

const Today = "\\2022\\Day14.txt"
const AoC2022* = "C:\\Users\\slayc\\source\\repos\\AdventOfCode\\RawData"
const SAND : string = "o"
const WALL : string = "#"

type
  State = object
    data                        : Table[string,string]
    sand                        : seq[string]  
    MinX                        : int
    MaxX                        : int  
    MaxY                        : int                

var s : State

proc ToCoord( x,y :int): string =
    $x & "," & $y


proc InAbyss( ipX , ipY :int): bool =
    
    if ipY >= s.MaxY:
        return true
  
    if ipX <= s.MinX:
        return true

    if ipX >= s.MaxX :
        return true

    return false
    

proc AddVerticalLine( ipX  ,  ipfromY ,   ipToY : var int) =
    var myStep = if (ipfromY < ipToY):  1 else: -1
    if myStep == -1:
        var myTmp = ipFromY
        ipFromY = ipToY
        ipToY = myTmp
        myStep = -myStep
    for myY in countup(ipFromY , ipToY):
        if not s.data.hasKey(ToCoord(ipX, myY)) :
            s.data[ToCoord(ipX, myY)] =WALL

   
proc AddHorizontalLine( ipFromX , ipTox ,  ipY: var int) =

    var myStep =if (ipfromX < iptox) : 1 else: -1
    if myStep == -1:
        var myTmp = ipFromX
        ipFromX=ipToX
        ipToX=myTmp
        myStep = -myStep
    for myX in countup(ipFromX ,ipTox):
        if not s.data.hasKey(ToCoord(myX, ipY)):
            s.data[ToCoord(myX, ipY)] = WALL
    


proc initialise() =
    s = State(
        #data : initTable[string, int](),
        sand : @[],
    )
    var RawData: seq[seq[seq[string]]] = (AoC2022 & Today).lines.toseq.mapIt(it.split(" -> ").toseq.mapit(it.split(","))) #.toseq.mapIt(x =x.parseInt)))
    for line in Rawdata:
        var myPoints=line
        for i in countup(1,myPoints.len-1):
            var myCurrentPoint =line[i]
            var myPrevPoint = line[i-1]
            var myFromX = parseInt(myPrevPoint[0])
            var myFromY = parseInt(myPrevPoint[1])
            var myToX = parseInt(myCurrentPoint[0])
            var myToY = parseInt(myCurrentPoint[1])

            s.MinX = min(myFromX,s.MinX)
            s.MinX = min(myToX,s.MinX)

            s.MaxX = max(myfromX,s.MaxX)
            s.MaxX = max(myToX,s.MaxX)

            s.MaxY = max(myFromY,s.MaxY)
            s.MaxY = max(myToY, s.MaxY)

            if myFromX == myToX:
                AddVerticalLine(myFromX, myFromY, myToY)

            if myFromY == myToY:
                AddHorizontalLine(myFromX, myToX, myFromY)

                

proc part01() =
    
    initialise()
        
    var myFallX: int
    var myFallY :int
    #Set s.Sand = Seq.Deb

    while true:
        myFallX = 500
        myFallY = 0
        
        while true:
            
            if InAbyss(myFallX, myFallY):
                break
        
            if not s.data.hasKey(ToCoord(myFallX, myFallY + 1)):
                myFallY += 1
                
            elif not s.data.hasKey(ToCoord(myFallX - 1, myFallY + 1)):
                myFallY += 1
                myFallX -= 1
                
            elif not s.data.hasKey(ToCoord(myFallX + 1, myFallY + 1)):
                myFallY += 1
                myFallX += 1
                
            else:
                s.data[ToCoord(myFallX, myFallY)] = SAND
                s.sand.add(ToCoord(myFallX, myFallY))
                break
        
        if InAbyss(myFallX, myFallY):
            break

    var myResult = s.sand.len
    
    var myOutput = "The answer to Day" & $Today[11..12] & "Part 01 is " & "768" &  " Found is " & $myResult
    echo myOutput


proc part02() =
   
    initialise()
    var myFallX :int
    var myFallY :int
    s.MaxY += 1
    
    while true:
        myFallX = 500
        myFallY = 0
        
        while true:
            
            if myFallY == s.MaxY:
        
                s.data[ToCoord(myFallX, myFallY)] = SAND
                s.sand.add(ToCoord(myFallX, myFallY))
                break
                
            elif not s.data.hasKey(ToCoord(myFallX, myFallY + 1)):
                
                myFallY += 1
                
            elif not s.data.hasKey(ToCoord(myFallX - 1, myFallY + 1)):
            
                myFallY += 1
                myFallX -= 1
                
            elif not s.data.hasKey(ToCoord(myFallX + 1, myFallY + 1)):
            
                myFallY += 1
                myFallX += 1
                
            elif (myFallX == 500) and (myFallY == 0):
            
                s.data[ToCoord(myFallX, myFallY)] = SAND
                s.sand.add( ToCoord(myFallX, myFallY))
            
                break
                
            else:
            
                s.data[ToCoord(myFallX, myFallY)] = SAND
                s.sand.add( ToCoord(myFallX, myFallY))
                
                break
                
            
            
       # Loop # Until myFallX = 500 And myFallY = 0
        if (myFallX == 500) and (myFallY == 0):
            break
    #Loop Until myFallX = 500 And myFallY = 0
        
    var myResult :int = s.sand.len
    var myOutput = "The answer to Day" & $Today[11..12] & "Part 02 is " & "26686" &  " Found is " & $myResult
    echo myOutput     


proc execute*()=
    part01()
    part02()