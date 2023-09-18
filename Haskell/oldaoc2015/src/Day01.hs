module Day01 (execute) where

import Constants 

-- Local constants
up = '('

down = ')'

today= "\\Day01.txt"

execute = do
    part01
    

part01 :: IO ()
part01 = do 
    moves <- readFile (Constants.rawDataPath  ++ today)
    let myFloor = moveDiff up down moves
    print( "The answer to our Day 01 Part 01 is 74. Found is " ++ show myFloor)

-- Calculate the difference between the up '(' and down ')' movements
moveDiff  :: Char->Char->[Char]->Int
moveDiff u d ms =  length [x | x<-ms, x == u] - length [ x| x<- ms, x == d]


-- part02 = do
--     xs <- readFile rawDataPath ++ year ++ day
--     let myFloor = count up xs - count down xs
--     print myFloor

-- part02 = do
--   xs <- readFile "C:\\Users\\slayc\\source\\repos\\AdventOfCode\\RawData\\2015\\Day01.txt"
  

-- code below works but recommendations are to not use head and tail as this is not idiomatic Haskell.
--count  a b :: a-> [b] ->
-- count :: Char -> [Char] ->Int
-- count myChar myList
--   | null myList = 0
--   | myChar/=head myList = 0 + count myChar (tail myList)
--   | myChar==head myList = 1 + count myChar (tail myList)

