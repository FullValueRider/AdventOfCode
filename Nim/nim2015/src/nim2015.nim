import strformat
# import strutils
import times


# import Day01_2015
# import Day02_2015
# import Day03_2015
# import Day04_2015
# import Day05_2015
# import Day06_2016
# import Day10_2015
# import Day05
# import Day06
# import Day07

when isMainModule:
    var myTime:float = cpuTime()
  # Day01_2015.execute()
  # Day02_2015.execute()
  # Day03_2015.execute()
  

  # Day04_2015.execute()
  Day05_2015.execute()
  # Day06.Execute()
  # Day07.Execute()
  # Day10_2015.execute()

    var elapsedTime = (cputime()-myTime)*1000
    echo fmt"Time (ms)= {elapsedTime}"

