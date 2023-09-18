import strformat
import strutils
import sequtils
import tables

import ../../AoCLib/src/constants
import BulbDisplay
const TODAY                             : string = "\\Day06.txt"
const Year                              : string ="\\2015"

type 
    Instruction = tuple
        message: string
        subArea: Area

    State = object
        Display                                    :BulbDisplay
        Data                                       :seq[string]
        Instructions                               :seq[Instruction]
        Lights                                     :Area



var s                                              :State

proc initialise() =
    var myData = (AocData & Year & Today).lines.toSeq
   
    var s = State()
    s.Instructions = myData 
        .mapIt(multiReplace(it,(" through ", ","), ("turn ", ""), (",", " "), ("  ", " ")))
        .mapIt(split(it," ").toSeq)
        .mapIt( (message:it[0], subArea:(xOrg:it[1].parseInt, yOrg:it[2].parseInt, xMax:it[3].parseInt, yMax:it[4].parseInt)))
        



proc Part01()

    Initialise
    
    Dim myArea:Variant
    myArea = Array(0, 0, 999, 999)
    ReDim Preserve myArea(1 To 4)
    
    Set s.Display = BulbDisplay(myArea)
    s.Display.SwitchOff myArea
    
    Dim myItem:Variant
    Dim myItems:Iteritems: Set myItems = Iteritems(s.Instructions)
    Do
        Set myItem = myItems.curitem(0)
        Dim myS:seqC
        Set myS = myItem
    
        Select Case myS.First

            Case "on": s.Display.SwitchOn myS.Tail.Toarray

            Case "off":  s.Display.SwitchOff myS.Tail.Toarray

            Case "toggle": s.Display.Toggle myS.Tail.Toarray

        End Select

    Loop While myItems.MoveNext
    
    Dim myResult:Long: myResult = s.Display.LitBulbs(myArea)
    
    fmt.dbg "The answer for Day {0} Part 01 is 543903. Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult




proc Part02()

    Initialise
    
    Dim myArea:Variant
    myArea = Array(0, 0, 999, 999)
    ReDim Preserve myArea(1 To 4)
    
    Set s.Display = BulbDisplay(myArea)
    s.Display.SwitchOff myArea
    s.Display.UseBrightness = True
    
    Dim myItem:Variant
    Dim myItems:Iteritems: Set myItems = Iteritems(s.Instructions)
    Do
        Set myItem = myItems.curitem(0)
        Dim myS:seqC
        Set myS = myItem
    
        Select Case myS.First

            Case "on": s.Display.SwitchOn myS.Tail.Toarray

            Case "off":  s.Display.SwitchOff myS.Tail.Toarray

            Case "toggle": s.Display.Toggle myS.Tail.Toarray

        End Select

    Loop While myItems.MoveNext
    
    Dim myResult:Long: myResult = s.Display.Brightness(myArea)
  
    fmt.dbg "The answer for Day {0} Part 02 is 14687245. Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult





   
proc Execute()
    Part01
    Part02
