module Main (main) where

import Day01 ( execute )
import Day02 (execute)

main :: IO ()
main = do
  Day01.execute
  Day02.execute
  print "Finished"
