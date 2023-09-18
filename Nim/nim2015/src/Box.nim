import strutils
import sequtils


const
    SEPARATOR                       = "x"
    
type 
    Edge = enum
        IsLength = 0,
        IsWidth = 1,
        IsHeight = 2

    Properties = object 
        dims: seq[int]

    State = object   
        areas: seq[int]
        perims: seq[int]

    Box = object  
        s: State
        p: Properties

var s*: State
var p*: Properties

proc newBox*(ipDims :string): Box =
    var myP = Properties(dims:ipDims.split(SEPARATOR).mapIt( parseInt(it)))

    var myState = State(areas: 
        [
            myP.dims[IsLength.ord] * myP.dims[IsHeight.ord], myP.dims[IsLength.ord] * myP.dims[IsWidth.ord], myP.dims[IsWidth.ord] * myP.dims[IsHeight.ord]
        ].toSeq,

        perims:
        [
             2 * (myP.dims[IsLength.ord] + myP.dims[IsHeight.ord]), 2 * (myP.dims[IsLength.ord] + myP.dims[IsWidth.ord]), 2 * (myP.dims[IsWidth.ord] + myP.dims[IsHeight.ord])
        ].toSeq)
    
    result = Box(p:myP, s:myState)
  
proc surfaceArea*(me:  Box): int =
    result = me.s.areas.foldl(a + b)*2
    

proc volume*(me: Box): int =
    result = me.p.dims[IsLength.ord] * me.p.dims[IsWidth.ord] * me.p.dims[IsHeight.ord]
    

proc areaOfSmallestFace*(me: Box): int =
    result = me.s.areas.min 


proc smallestPerimeter*(me:Box): int =
    result = me.s.perims.min


proc wrappingSize*(me: Box): int =
    result = me.surfacearea + me.areaOfSmallestFace  


proc ribbonLength*(me: Box): int =
    result = me.smallestPerimeter + me.volume



