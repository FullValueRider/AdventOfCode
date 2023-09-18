module Day02 (execute) where

import Constants 
import Data.List.Split (splitOn)

-- Local constants
today :: String
today= "\\Day02.txt"

execute :: IO ()
execute = do
    part01
   -- part02
    

part01 :: IO ()
part01 = do 
    moves <-  readFile (Constants.rawDataPath  ++ today)
    let myLines = lines moves
    let myBoxes = [splitOn "x" x| x<- myLines]
    print myBoxes
    
    
    --print( "The answer to Day 01 Part 01 is 74. Found is " ++ show myFloor)



-- part02 :: IO ()
-- part02  = do
--     moves <- readFile (Constants.rawDataPath  ++ today)
--     let myFloor  = elemIndex (-1) (scanl (+) 0 [moveToInt x| x<-moves])
--     print( "The answer to Day 01 Part 02 is 1795. Found is " ++ show myFloor)

-- moveToInt :: Char -> Int
-- moveToInt x
--     | x==up = 1
--     | x== down = -1
--     | otherwise = 0
