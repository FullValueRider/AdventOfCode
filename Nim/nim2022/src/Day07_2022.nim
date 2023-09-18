import sequtils
import table
#========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

    # Day 07 is based on binary trees.
const Today            : string = "\\2022\\Day07.txt"
const IS_COMMAND       : string = "$"

type 
    State = object 
        Data                                    : seq

var s                          : State

type
    Command = enum   

        Type = 1
        Action = 2
        Target = 3



proc calculateDirectorySizes(ipFS: ElfComDir): Table

    var mysizes: Lookup = Lookup.Deb
    ipFS.cd("/")
    ipFS.Size mysizes
    return mysizes



proc Part01*()

Initialise
var myFS: ElfComDir = constructFileSystem(s.Data)
var myDirectorySizes: Lookup = CalculateDirectorySizes(myFS)
var myResult: int = myDirectorySizes.ReduceIt(rdSum(cmpLTEQ(100000)))

echo fmt Fmt.Text("The answer to Day {0} part 1 is {1}.  Found is {2}", Mid$(Today, 10, 2), "1792222", myResult)




proc Part02*()
# The only change from part 1 was increasing the count from 4 to 14
Initialise

Initialise
var myFS: ElfComDir = constructFileSystem(s.Data)
var myDirectorySizes: Lookup = CalculateDirectorySizes(myFS)

var myUsedSpace: int = myDirectorySizes.SortByItem.Last.Item(0)
var myUnusedSpace: int = 70000000 - myUsedSpace
var myMinDeletion: int = 30000000 - myUnusedSpace
var myResult: int = myDirectorySizes.FilterIt(cmpMT.Deb(myMinDeletion)).SortByItem.First.Item(0) #.sort.first
        
echo fmt "The answer to Day {0} part 2 is {1}.  Found is {2}", Mid$(Today, 10, 2), "1112963", myResult



proc Initialise*()
    s.Data = seq.Deb(Filer.GetFileAsArrayOfstrings(AoC & Today, vbCrLf)) _
    .MapIt(mpSplitToSubStr(ToSeq, Char.twSpace))


proc constructFileSystem*(ipLog: var seq ): ElfComDir

var myFS: ElfComDir = ElfComDir.Deb("/")

var myLogItems: IterItems = IterItems.Deb(ipLog)
while
    
                if LogItemIsCommand(myLogItems.Item(0)) :
        
        var myCommand: seq = seq.Deb
        while
            myCommand.Add myLogItems.Item(0)
            
            if myLogItems.Has

                if LogItemIsCommand(myLogItems.Item(1)) :
                    break
                
            
            
        while  myLogItems.Move #Check for top of loop 

        
        updateFS myFS, myCommand
        
    
    
while  myLogItems.Move #Check for top of loop 


var myreturn: ElfComDir = myFS.cd("/")
return myreturn



proc LogItemIsCommand*(ipLogItem: var seq ): bool

return ipLogItem.Item(1) = "$"



proc updateFS*(ipFS: var ElfComDir , ipCommand: var seq )


case ipCommand(1)(Action)  #Note: Nim requires constant expressions for each of
    
    of "cd"
        
        var myDir: string = $ipCommand(1)(Target))
            ipFS = ipFS.cd(myDir)
    
    of "ls" :
    
        var myItems: IterItems = IterItems(ipCommand).SetFTS(1)
        while
            var myItem: seq = myItems.Item(0)
            Kase myItem(1)  #Note :: Nim requires constant expressions for each of
            
                of "dir"
                    var myNewDir: string = myItem(2)
                    if ipFS.Dirs.LacksItem(myNewDir) :
                        ipFS.Dirs.Add myNewDir, ElfComDir.Deb(myNewDir, ipFS)
                    
            
                else ::
                
                    ipFS.Files.Add myItem(2), parseInt(myItem(1))
                    
            
            
        while  myItems.Move #Check for top of loop 

        






    
proc execute*() = 
    part01()
    part02()
