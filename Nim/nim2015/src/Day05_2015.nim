import strutils
import sequtils
import strformat
import ../../AocLib/chars
import ../../AoCLib/constants

const 
    Today                       = "\\Day05.txt"
    Year                        = "\\2015"
   


type
  State = object
    badSubstrings               : seq[string]
    vowels                      : seq[char]
    data                        : seq[string]
   
var s : State

proc hasBadSubstrings( ipWord : string) : bool =

    for myBad in s.badSubstrings:
        if ipWord.contains(myBad):
            return true

    return false

proc hasLessThanThreeVowels(ipWord: string) : bool =
    var mycount:int =0
    for myChar in ipWord:
        if s.vowels.contains(myChar):
            mycount+=1
    return if myCount<3: true else: false
   

proc hasNoDoubleChars( ipWord : string) : bool =
    
    for myIndex in countup(0,ipword.len-2):
        if ipWord[myIndex] == ipWord[myIndex+1]:
            return false
        
    return true

proc isNiceV1( ipWord : string) : bool =
    #echo ipWord
    if hasBadSubstrings(ipWord):
        #echo "hasBadSubStrings";
        return false
    if hasNoDoubleChars(ipWord):
        #echo "has no double chars";
        return false
    if hasLessThanThreeVowels(ipWord):
        #echo "has less then three vowels";
        return false
    #echo $true;
    return true


proc hasMultiplePairs( ipWord : string) : bool =

    # The word needs to have at least 4 characters to have double pairs
    if ipWord.len < 4:
        return false
    
    for myIndex in countup(0, ipWord.len-4):
       
        var myPair : string = ipWord[myIndex] & ipWord[myIndex+1]
        if (ipWord.len - ipWord.replace( myPair, chars.twNoString).len)  > 3 :
            return true

    return false


proc lacksMultiplePairs( ipWord : string) : bool =
    return if hasMultiplePairs(ipWord)==true:  false  else: true


proc hasSpacedRepeats( ipWord : string) : bool =

    for myIndex in countup(0,ipWord.len - 3):
    
        if ipWord[myIndex] == ipWord[myIndex + 2] :
            return true
    
    return false


proc lacksSpacedRepeats(ipWord : string) : bool =
    return if hasSpacedRepeats(ipWord): false else: true


proc isNiceV2( ipWord : string) : bool =

    if lacksMultiplePairs(ipWord):
        return false
    if lacksSpacedRepeats(ipWord):
        return false
    return true


proc initialise() =
    s=State(
        data : @[],
        badSubstrings : @["ab", "cd", "pq", "xy"],
        vowels : @['a', 'e', 'i', 'o', 'u'])
        
    s.data = (AocData & Year & Today).lines.toseq

proc part01() =
    
    initialise()
    var myResult:int = s.data.mapIt(isNiceV1(it)).countIt(it == true)
    # for myWord in s.data:

    #     if isNiceV1(myWord):
    #         myResult += 1
    echo fmt"The answer to Day {Today[5..6]} part 01 is 238 .  Found is {myResult}"


proc part02() =
   
    initialise()
    var myResult:int = s.data.mapIt(isNiceV2(it)).countIt(it == true)
    # var myResult : int
    # for myWord in s.data:
    #     if isNiceV2(myWord):
    #         myResult += 1

    echo fmt"The answer to Day {Today[5..6]} part 2 is 69.  Found is {myResult}"        


proc execute*()=
    part01()
    part02()