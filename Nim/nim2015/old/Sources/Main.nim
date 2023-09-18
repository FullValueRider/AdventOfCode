import strformat
import strutils
import times
import Day01


# import Day02
# import Day03
# import Day04
# import Day05
# import Day06
# import Day07

var myTime:float = cpuTime()
Day01.Execute()
# Day02.Execute()
# Day03.Execute()
# Day04.Execute()
# # #Day04Tables.Execute()
# Day05.Execute()
# Day06.Execute()
# Day07.Execute()

echo fmt"Time (ms)= {(cputime()-myTime)*100}"