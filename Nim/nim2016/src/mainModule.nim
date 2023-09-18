import strformat
import strutils
import times


import Day01_2016
# import Day02_2015
# import Day03_2015
# import Day04_2015


when isMainModule:
  var myTime:float = cpuTime()
  Day01_2016.execute()
  # Day02_2015.execute()
  # Day03_2015.execute()
  # Day04_2015.execute()
  # Day05_2015.execute()
  # Day06.Execute()
  # Day07.Execute()

  echo fmt"Time (ms)= {(cputime()-myTime)*100}"

