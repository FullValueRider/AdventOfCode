import sequtils
import tables


type

    State = ref object
        Numbers: Table[int,int]
        HRanks:   seq[Table[int,int]]
        VRanks:   seq[Table[int,int]] 
        Index: int

    BoardObj* = object
        s: State

    
var Index*:int
    

proc newBoard*( ipBoardNumbers: seq[seq[int]]):  BoardObj =
    var s= State(
        Numbers:Table[int,int](),
        HRanks: @[], #seq[Table[int,int]]() #newseqwith(ipboardnumbers.len(), initTable[int,int]()),
        VRanks: @[] # newseqwith(ipBoardNumbers.len(), initTable[int,int]())
    )

    # Use the index field in this file as a counter
    s.Index = Board.Index
    Board.Index += 1
    
    #Make a list of the numbers on the Board, there are no duplicated numbers
    for myRowNumbers in ipBoardNumbers:
        for myNumber in myRowNumbers:
            s.Numbers[mynumber]=0

    # make a list of the horizontal and vertical ranks on the board
    for myRownumbers in ipBoardNumbers:
        # nim can be very frustrating at times
        # I couldn't work out the syntax for adding two sequences to a Table
        # var myZip = zip(myrownumbers,repeat(0.int,myrownumbers.len()))
        var myTable=Table[int,int]()
        for myNumber in myRowNumbers:
             myTable[myNumber]=0
        s.HRanks.add(myTable)

    for myIndex in 0..ipboardnumbers[0].high:
        var myTable=Table[int,int]()
        for myNumbers in ipBoardnumbers:
            myTable[myIndex]=0
        s.VRanks.add(myTable)

    var myBoard:BoardObj = BoardObj(s:s)
    myBoard.s = s
    return myboard

proc IsWinner*(me:BoardObj):bool  =
    for myHRank in me.s.Hranks:
        if myHRank.values.toseq().countit(it >= 0)==myHrank.len():
            return true
    for myVRank in me.s.Vranks:
        if myVRank.values.toseq().countit(it >= 0)==myVrank.len():
            return true


proc HasNumber*(me:BoardObj, ipNumber:int):bool =
    
    if ipNumber notin me.s.Numbers.keys.toSeq :
        return false
    
    var myHit:bool
    for myRank in me.s.HRanks:
        if ipNumber in myRank.keys.toseq:
            myRank[ipnumber] = 1
            myHit=true

    for myRank in me.s.Vranks:
        if ipNumber in myRank.keys.toseq:
            myRank[ipnumber] = 1
            myHit=true