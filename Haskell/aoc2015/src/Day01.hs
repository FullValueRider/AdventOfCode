module Day01 (execute) where

import Constants 
import Data.List ( elemIndex )

-- Local constants
up :: Char
up = '('

down :: Char
down = ')'

today :: String
today= "\\Day01.txt"

execute :: IO ()
execute = do
    part01
    part02
    

part01 :: IO ()
part01 = do 
    moves <- readFile (Constants.rawDataPath  ++ today)
    let myFloor  = sum [moveToInt x| x<-moves]
    -- let myFloor = moveDiff up down moves
    print( "The answer to Day 01 Part 01 is 74. Found is " ++ show myFloor)



part02 :: IO ()
part02  = do
    moves <- readFile (Constants.rawDataPath  ++ today)
    let myFloor  = elemIndex (-1) (scanl (+) 0 [moveToInt x| x<-moves])
    print( "The answer to Day 01 Part 02 is 1795. Found is " ++ show myFloor)

moveToInt :: Char -> Int
moveToInt x
    | x==up = 1
    | x== down = -1
    | otherwise = 0

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

