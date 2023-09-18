import strutils
import strformat
import sequtils

type

    VentLineType* = enum
        vlPoint,
        vlHorizontal,
        vlVertical,
        vlDiagonal,

    CoordType=enum
    
        ctX1,
        ctY1,
        ctX2,
        ctY2,

    State =object
        X1: int
        X2: int
        Y1: int 
        Y2: int
        IsVentLineType: VentLineType

    VentLineStartEndCoords* = object
        s: State

var me: VentLineStartEndCoords

proc initVentLineStartEndCoords*( ipCoords: string): VentLineStartEndCoords =
    var myCoords: seq[int] = ipCoords.split(",").mapit( it.parseint )
    me = VentLineStartEndCoords( 
        s: State(         
            X1 : mycoords[ctX1.int],
            Y1 : mycoords[ctY1.int],
            X2 : mycoords[ctX2.int],
            Y2 : mycoords[ctY2.int],
        ))
            
    # Determine any equivalence between X1,X2 and Y1,Y2 so that we
    # can assign a type to the line
    
    if ((me.s.X1 == me.s.X2) and (me.s.Y1 == me.s.Y2)):            
        me.s.IsVentLineType = VentLineType.vlPoint
    elif (me.s.Y1 == me.s.Y2):                                
         me.s.IsVentLineType = VentLineType.vlHorizontal
    elif (me.s.X1 == me.s.X2):                               
        me.s.IsVentLineType = VentLineType.vlVertical
    else:                                             
        me.s.IsVentLineType = VentLineType.vlDiagonal
            
    return  me


proc GetLineType*(me: VentLineStartEndCoords ): VentLineType =
        return me.s.IsVentLineType


iterator dodgyStepping(a, b:int):int =
    let step = if a < b: 1 else: -1
    var i = a
    while (if 1 == step: i <= b else: i >= b):
        yield i
        i += step   


proc GetVentLineCoords*(me: VentLineStartEndCoords): seq[string] =
        
    var myVentLine:seq[string] = @[]
 
    case  me.s.IsVentLineType:

        of vlPoint:
            myVentLine.add( fmt"{me.s.X1},{me.s.X2}"  )  #fmt("{me.s.X1},{me.S.Y1}")
            
        of vlHorizontal:
            for myXCoord in dodgyStepping(me.s.X1, me.s.X2):
                myVentLine.add( fmt"{myXCoord},{me.s.Y1}")
            
        of vlVertical:
            for myYCoord in dodgyStepping(me.s.Y1, me.s.Y2):
                myVentLine.add( fmt"{me.s.X1},{myYCoord}" )
        
        of vlDiagonal:
            var myYCoord = me.s.Y1
            var myYStep = if (me.s.Y1 <= me.s.Y2): 1 else: -1
            for myXCoord in dodgyStepping( me.s.X1, me.s.X2):
                myVentLine.add( fmt("{myXCoord},{ myYCoord}"))
                myYCoord += myYStep

    return myVentLine
        
