import times

#import Day01_2022
#import Day02_2022
#import Day03_2022
#import Day04_2022
#import Day05_2022
import Day06_2022
#import Day12_2022
#import Day14_2022
#import Day18_2022
#import Day21_2022
#import Day23_2022
#import Day24_2022



when isMainModule:
  var myTime:float = cpuTime()

  #Day01_2022.execute()
  #Day02_2022.execute()
  #Day03_2022.execute()
  #Day04_2022.execute()
  #Day05_2022.execute()
  #Day12_2022.execute()
  Day06_2022.execute()
  #Day14_2022.execute()
  #Day18_2022.execute()
  #Day21_2022.execute()
  #Day23_2022.execute()
  #Day24_2022.Execute()
  

  var myFinishTime :float = (cputime()-myTime) * 1000
  echo "Time (ms)= "  & $myFinishTime
