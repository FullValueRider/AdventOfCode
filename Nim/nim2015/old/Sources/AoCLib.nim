const AocData*           :string = "C:\\Users\\slayc\\source\\repos\\AdventOfCode\\RawData"
# type
#   IComparer* = object
#     execCmp* : proc(x:int,y:int):bool
    



#  proc transpose*[T](ipSource:seq[seq[T]]): seq[seq[T]]=
#     var myTransposed:seq[seq[T]]=newseq[seq[T]](ipsource[0].len())
#     for myItem in ipSource:
#          for myIndex in 0..myItem.high:
#             myTransposed[myIndex].add myItem[myIndex]
#      return myTransposed


# proc cmpEq*[T](testItem:T, refItem:T) :bool= testitem == refItem
# proc cmpNEq*[T](testItem:T, refItem:T):bool = testitem != refItem
# proc cmpMT*[T](testItem:T, refItem:T):bool = testitem > refItem
# proc cmpMTEq*[T](testItem:T, refItem:T):bool = testitem >= refItem
# proc cmpLT*[T](testItem:T, refItem:T):bool = testitem < refItem
# proc cmpLTEq*[T](testItem:T,refItem:T):bool = testitem <= refItem


# proc toString*(str: seq[char]): string =
#   result = newStringOfCap(len(str))
#   for ch in str:
#     add(result, ch)
