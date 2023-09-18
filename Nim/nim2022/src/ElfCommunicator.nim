import tables
#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C


#    const Today            : string = "\2022\DayXXTest1.txt"


type 

    State = object
        Ticker                                  : int
        Xreg                                    : int
        Col                                     : int
        Row                                     : int
        
    Properties = object 
        Program                             : seq
        Result                              : Table
        VDU                                 : seq

var    s                                        : State

    type 

        
            


    p                          : Properties

proc Deb*(): ElfCommunicator = 
    With New ElfCommunicator
        return .constructInstance
    


    proc constructInstance*(): ElfCommunicator
    return Me


proc  Program*(me: ElfCommunicator): seq = 
    return me.p.Program


proc  VDU*(): seq = 
    return p.VDU


    Property  Program(ipProgram: var seq )
        p.Program = ipProgram.Clone


proc Result*(): Lookup = 
    return p.Result


proc Run*()     = 
    s.Xreg = 1
    s.Ticker = 0
        p.Result = Lookup.Deb
        p.VDU = seq.Deb.Repeat(string(40, Char.twPeriod), 6).MapIt(mpSplitToChars(ToSeq))
    
    var myInstructions: IterItems = IterItems.Deb(p.Program)
    while
        
        var myOp: string = myInstructions.Item(0)(1)
        
        case myOp  #Note: Nim requires constant expressions for each of
        
            of "noop" :
            
                UpdateVDU
                s.Ticker += 1
                CheckForOutput
                
                
            of "addx" :
            
                var myNumber: int = myInstructions.Item(0)(2)
                var myTick: int
                For myTick = 1 To 2
                    UpdateVDU
                    s.Ticker += 1
                    CheckForOutput
                    
                

                s.Xreg += myNumber
                
        
        
    while  myInstructions.Move #Check for top of loop 

    



    proc CheckForOutput*()

    if s.Ticker = 20 :
        p.Result.Add s.Ticker, s.Xreg
    else:if (s.Ticker - 20) Mod 40 = 0 :
        p.Result.Add s.Ticker, s.Xreg
    
    



    proc UpdateVDU*()
    
    s.Row = s.Ticker \ 40
    s.Row = s.Row Mod 6
    s.Col = s.Ticker Mod 40
    
    # seq indexing starts at 1 so we need to add 1 to get the correct row and column addresses
    # when accessing VDU
    if s.Col >= s.Xreg - 1 and s.Col <= s.Xreg + 1 :
        p.VDU.Item(s.Row + 1).Item(s.Col + 1) = "#"
    else:
        p.VDU.Item(s.Row + 1).Item(s.Col + 1) = "."