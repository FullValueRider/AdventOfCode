import mytables
import sequtils
#=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

const ROOT                                      : string = "/"
const NO_NAME                                   : string = ""

type 

    FileInfo = enum 
        Size = 1
        Name = 2

    Properties = ref object 
    
        Parent                              : ElfComDir
        Name                                : string
        Directories                         : OrderedTable[string, ElfComDir] #Lookup # name vs elfcomdir
        Files                               : OrderedTable[string, int] #Lookup # name vs size
        Path                                : string
    
    ElfComDir = object
        p                                   : Properties


# var     p                                   : Properties
    

proc newElfComDir*(ipName:  string ) : ElfComDir =
    var myP :Properties =
        Properties(
            Name : ipName,        
            Directories :initOrderedTable[string, ElfComDir](),
            Files : initOrderedTable[string, int](), 
            Path : ""
        )
    var myElfcomDir = ElfComDir(
        p: myP)
       
    return myElfComDir

proc newElfComDir*(ipName:  string , ipParent : ElfComDir) : ElfComDir =
    var myP :Properties =
        Properties(
            Name : ipName,  
            Parent: ipParent,      
            Directories :initOrderedTable[string, ElfComDir](),
            Files : initOrderedTable[string, int](), 
            Path : ""
        )
    var myElfcomDir = ElfComDir(
        p: myP)
       
    return myElfComDir
         
    
    
proc  Name*(me: ElfComDir): string = 
    return me.p.Name
    
    
proc  `Name =`*( me: ElfComDir, ipName: var string ) = 
    me.p.Name = ipName


proc  Dirs*(me: ElfComDir): OrderedTable[string, ElfComDir] = 
    return me.p.Directories


proc `Dirs =`*(me : ElfComDir, ipDirectories: OrderedTable[string, ElfComDir] ) =
        me.p.Directories = ipDirectories
    
    
    
proc  Files*(me : ElfComDir): OrderedTable[string, int] = 
    return me.p.Files
    
        
proc `Files =`*(me : ElfComDir, ipFiles: OrderedTable[string, int] ) =
         me.p.Files = ipFiles
    
    
    
proc  Parent*(me: ElfComDir): ElfComDir = 
        return me.p.Parent
    
    
proc `Parent =`*(me: ElfComDir, ipParent: var ElfComDir ) =
         me.p.Parent = ipParent
    
    
proc cd*(me: ElfComDir, ipName: string ): ElfComDir = 

        case ipName  #Note: Nim requires constant expressions for each of
    
            of ".." :
            
                if me.p.Name == ROOT :
                    return me
                else:
                    return me.p.Parent
                
                
            of ROOT :
            
                #echo fmt Name, ROOT
                if me.Name == ROOT :
                    return me
                
                var myDir: ElfComDir = me
            
                while  myDir.Name != ROOT:
                        myDir = myDir.Parent
                
                
                return myDir
                
                
            else :

                return  me.p.Directories[ipName]
        
    
proc Size*(me: ElfComDir, ipDirSizes: var OrderedTable[string, int] ) = 
       
        var mysize: int = 0
        
        if me.p.Directories.len > 0 :
            for myDir in me.p.Directories.values:
            
                if myDir.Name == NO_NAME : 
                    continue

                if me.p.Directories.len == 0 :
                    continue
                
                mydir.Size ipDirSizes
                mysize += ipDirSizes.last[0]
                
        if me.p.Files.len > 0 :
            mysize += me.p.Files.values.toSeq.foldl(a + b)
        
        var myName: string = me.Name
        # Make the name unique
        while  ipDirSizes.hasKey(myName):
            myName &= "_"
       
        ipDirSizes[myName] = mysize
     
    
    

