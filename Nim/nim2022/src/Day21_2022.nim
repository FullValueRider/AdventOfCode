import strutils
import sequtils
#import strformat
import tables
import parseutils
import ..\..\AoCLib\chars

const AoC2022                   = "C:\\Users\\slayc\\source\\repos\\AdventOfCode\\RawData"
const Today                     = "\\2022\\Day21.txt"
    
const ADD                       : string = "+"
const SUB                       : string = "-"
const MUL                       : string = "*"
const DIV                       : string = "/"

type 
    State = object
        Numbers                 : Table[string,string]
        Formulas                : Table[string,seq[string]]
    
var s                           : State

proc dequeue[T]( ipSeq : var seq[T]):  T =
    result = ipSeq[0]
    ipSeq.delete(0..0)
    

proc holdsKey[K,T]( ipTable: Table[K,T], ipKey: K) : bool =
    return ipTable.haskey(ipKey)

proc lacksKey[K,T]( ipTable: Table[K,T], ipKey: K) : bool =
    return not ipTable.hasKey(ipKey)

proc isRoot(ipKey:string ): bool =
    return ipKey == "root"
    

proc isHumn(ipKey: string ): bool=
    return ipKey == "humn"

proc isNumeric(ipString: string) : bool =
    var tmp:int
    return ipString.parseSaturatedNatural(tmp) == 0


proc updateRemainingFormulasWithNumbers(ipFormulas: var Table[string,seq[string]] ) =
    
    for (myKey,myFormula) in ipFormulas.mpairs:
        var myM1: string = myFormula[0]
        var myM2: string = myFormula[2]
        
        if isHumn(myM1) : 
            continue
        if s.Numbers.hasKey(myM1) : 
            myFormula[0] = s.Numbers[myM1]
        
        myM2 = myFormula[2]
        if isHumn(myM2) : 
            continue
        if s.Numbers.hasKey(myM2) : 
            myFormula[2] = s.Numbers[myM2]
  

proc rewriteFormulas(): Table[string,seq[string]] =

    var myNewFormulas: Table[string,seq[string]]
    for (myKey,myFormula) in s.Formulas.mpairs:
   
        var myKey: string = myKey
        var myM1 = myFormula[0]
        var myM2 = myFormula[1]
        var myOp = myFormula[2]
        
        case myOp  #Note: Nim requires constant expressions for each of
        
            of ADD :
            
                if myM1.isNumeric :   #.IsNumeric :
                    myNewFormulas[myM2] = @[myKey, SUB, myM1]
                else:
                    myNewFormulas[myM1]= @[myKey, SUB, myM2]
                
                
            of SUB :
            
                if myM1.isNumeric :
                    myNewFormulas[myM2] = @[myM1, SUB, myKey]
                else:
                    myNewFormulas[myM1] = @[myKey, ADD, myM2]
                
            
            of MUL :
            
                if myM1.isNumeric :
                    myNewFormulas[ myM2] = @[myKey, DIV, myM1]
                else:
                    myNewFormulas[myM1] = @[myKey, DIV, myM2]
                
            
            of DIV :
        
                if myM1.isNumeric :
                    myNewFormulas[myM2] = @[myM1, DIV, myKey]  
                else:
                    myNewFormulas[myM1] = @[myKey, MUL, myM2]
    
    return myNewFormulas


proc findNextNewFormula*(ipKey: var string , ipformulas: var Table[string, seq[string]] ): (string, seq[string]) =
    
    for (myKey,myFormulas) in ipformulas.mpairs:
 
        var myFormula: seq[string] = myFormulas
        if myFormula[0] == ipKey :
            return (myKey, myFormula)
        
        if myFormula[2] == ipKey :
            return (myKey, myFormula)
    
    return ("", @[])
    


proc evaluateFormula(ipLHS, ipOp, ipRHS: string): string =
    var myLHS : int64 = ipLHS.parseInt
    var myRHS : int64 = ipRHS.parseInt
    case ipOp  #Note: Nim requires constant expressions for each of
        of ADD:       
            result = $(myLHS + myRHS)
        of MUL:       
            result = $(myLHS * myRHS)
        of SUB:       
            result = $(myLHS - myRHS)
        of DIV:       
            result = $(myLHS div myRHS)
        else:
            raise
    
     
proc initialise*() =

    var myData = (AoC2022 & Today).lines.toSeq
        .mapIt(it.replace(twColon,twNoString))
        .mapIt(it.split(twSpace))

    for myMonkey in myData.mitems:
        case myMonkey.len  #Note: Nim requires constant expressions for each of
            of 2:
                s.Numbers[myMonkey[0]]=myMonkey[1]
            else:
                s.Formulas[myMonkey.dequeue] = myMonkey
        
   

proc part01() =

    initialise()
    
    var myKey: string
    var myNumber: string

    while true:   
        var myNewNumbers: seq[string] = @[]
        for (myKey,myFormulas) in s.Formulas.mpairs:
           # var myFormulasCount = s.Formulas.len
            var myM1: string = myFormulas[0]
            if s.Numbers.lacksKey(myM1):
                continue
            
            var myM2: string = myFormulas[2]
            if s.Numbers.lacksKey(myM2) :
                continue
            
            var myLHS = s.Numbers[myM1]
            var myRHS = s.Numbers[myM2]
            var myOp = myFormulas[1]
            
            myNumber = evaluateFormula(myLHS, myOp, myRHS)
            
            myNewNumbers.add( myKey)
            s.Numbers[myKey] = myNumber
        
        if isRoot(myKey) :
                break
        
        for myRemoveNumber in myNewNumbers:
            s.Formulas.del( myRemoveNumber)
    
    var myResult: int64 = myNumber.parseInt
                                                                                                        
    #echo fmt "The answer to Day {Today[11..12} part 1 is 145167969204648.  Found is {myresult}"
    var myOutput = "The answer to Day" & $Today[11..12] & "Part 01 is " & "145167969204648" &  " Found is " & $myResult   
    echo myOutput


proc part02*() =

    initialise()
    
    while true:
        var myNewNumbers: seq[string] = @[]
        var myFormulasCount: int = s.Formulas.len
        var myKey : string
        for (myKey,myFormula) in s.Formulas.mpairs:
        
            if isHumn(myKey) :
                continue
            var myM1: string = myFormula[1]
            if s.Numbers.lacksKey(myM1) : 
                continue
            if isHumn(myM1) :
                continue
                
            var myM2: string = myFormula[3]
            if s.Numbers.lacksKey(myM2) :
                continue
            if isHumn(myM2) :
                continue
            
            var myLHS = s.Numbers[myM1]
            var myRHS = s.Numbers[myM2]
            var myOp = myFormula[2]
            
            var myNumber = evaluateFormula(myLHS, myOp, myRHS)
            
            
            myNewNumbers.add( myKey)
            s.Numbers[myKey] = myNumber

        if isRoot(myKey) :
            break
        if myNewNumbers.len == 0 :
            continue
        for myRemoveNumbers in myNewNumbers.items:
            s.Formulas.del( myRemoveNumbers)
        
        if s.Formulas.len == myFormulasCount:
            break 
    # we now have only formulas that depend on humn
    
    updateRemainingFormulasWithNumbers( s.Formulas)
    var myNewFormulas =  rewriteFormulas()
    
    
    # get the seed value for calculating humn
    var mySeed: int64 = 0
    var myOldKey: string = "root"
    if s.Formulas["root"][0].isNumeric :
        mySeed = s.Formulas["root"][0].parseInt
    else:
        mySeed = s.Formulas["root"][2].parseInt
    
    
    # now work through the rewritten formula until we find #humn#; the new key
    # the first new formula returned is that for root which we need to skip over
    #; need the next key to pair with the equality value in the root formula.
    # get the array containing key and new formula
    var myPackage = findNextNewFormula(myOldKey, myNewFormulas)
    var myNewKey = myPackage[0]
    var myFormula: seq[string]
    while true:
        myOldKey = myNewKey
        myPackage = findNextNewFormula(myOldKey, myNewFormulas)
        myNewKey = myPackage[0]
        var myformula = myPackage[1]
        
        if isHumn(myNewKey) :
            break
        
        if myformula[0] == myOldKey :
            mySeed = evaluateFormula($mySeed, myformula[2], myformula[3]).parseInt
        else:
            mySeed = evaluateFormula(myformula[1], myformula[2], $mySeed).parseInt
    
    # the formula for humn has two names, one of which will be translatable to a number
    # using s.numbers, the other is the name for myseed, but we don#t know which way around
    # the two names will be, but if we find a name in s.numbers we know the other is myseed
    if s.Numbers.holdsKey(myPackage[1][0]) :
        
        var tmp1 = s.Numbers[myPackage[1][0]]
        var tmp2 = myformula[1]
        var tmp3 =  $mySeed
        mySeed  = evaluateFormula(tmp1, tmp2,tmp3).parseInt
        
    else:
        
        mySeed = evaluateFormula($mySeed, myformula[1], s.Numbers[myPackage[1][2]]).parseInt
        
    var myResult: int64 = mySeed
    
    #echo fmt "The answer to Day {0} part 2 is {1}.  Found is {2}", Mid$(Today, 10, 2), "3330805295850", myResult
    var myOutput = "The answer to Day" & $Today[11..12] & "Part 02 is " & "145167969204648" &  " Found is " & $myResult   
    echo myOutput

proc execute*() =
    
    part01()
    #part02()
