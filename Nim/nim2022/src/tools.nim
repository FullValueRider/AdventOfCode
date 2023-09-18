

const FourwayOffsets*                           : seq[seq[int]] = @[ @[0,1], @[1,0], @[0,-1], @[-1,0]]
const EightWayOffsets*                          : seq[seq[int]] = @[ @[0,1], @[1,0], @[0,-1], @[-1,0], @[1,1], @[1,-1], @[-1,-1], @[-1,1]]


type 
    CompassPoints* = enum 
        Fourway
        EightWay


iterator reverse*[T](a: seq[T]): T {.inline.} =
    var i = a.high
    while i > -1:
        yield a[i]
        dec(i)
        

proc first*[T](ipSeq: seq[T]): T =
    return ipSeq[ipSeq.low]

proc last*[T](ipSeq: seq[T]): T =
    if ipSeq.len == 0:
        echo "sequence is empty"
    else:
        return ipSeq[ipSeq.high]

proc lhsOnly*[T](ipLhs: seq[T], ipRhs: seq[T]): seq[T] =
    for myItem in ipLhs:
        if myItem notin ipRhs:
            result.add(myItem)

proc rhsOnly*[T](ipLhs: seq[T], ipRhs: seq[T]): seq[T] =
    for myItem in ipRhs:
        if myItem notin ipLhs:
            result.add(myItem)

proc inBoth*[T](ipLhs:seq[T], ipRhs: seq[T]): seq[T] =
    for myItem in ipRhs:
        if myItem in ipLhs:
            result.add(myItem)

proc notInBoth*[T](ipLhs:seq[T], ipRhs: seq[T]): seq[T] =
    for myItem in ipRhs:
        if myItem notin ipLhs:
            result.add(myItem)
    for myItem in ipLhs:
        if myItem notin ipRhs:
            result.add(myItem)

proc unique*[T](ipLhs:seq[T], ipRhs: seq[T]): seq[T] =
    result = ipLhs
    for myItem in ipRhs:
        if myItem notin iplhs:
            result.add(myItem)

proc splitAt*[T](ipSeq: seq[T], ipSplitIndex: int): seq[seq[T]] =
    result.add(ipseq[0..ipSplitIndex-1])
    result.add(ipseq[ipSplitIndex .. ^1])

proc reverse*(s: var string) =
  for i in 0 .. (s.high div 2):
    swap(s[i], s[s.high - i])

proc holdsItem*[T]( ipSeq: seq[T], ipItem: T): bool =
    return ipSeq.contains(ipItem)

proc lacksItem*[T]( ipSeq:seq[T], ipItem: T): bool =
    return not ipSeq.contains(ipItem)



proc dequeue*[T](ipSeq: var seq[T]): T =
    result = ipSeq[0]
    ipSeq.delete(0)
    return result

proc indexOf*[T](ipSeq: seq[T], ipItem: T): int =
    for myIndex,myItem in ipSeq:
        if myItem == ipItem:
            return myIndex