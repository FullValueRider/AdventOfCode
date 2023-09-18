module Day01 (execute) where

import MyConstants

-- Local constansts
day :: String
day= "\\Day01.txt"

execute :: IO ()
execute = do
    part01
   -- part02
    
part01 :: IO ()
part01 = do 
    
    putStr "The answer for Day 01 Part 01 is 74. Found is " 
    print myDiff
    


-- walk :: (Int,Int) ->[Char]->(Int,Int)
-- walk (x,y) xs =
--     | fst x == R 
--     | fst xs == 'R'
--         | (x,y)==(0,1 ), (-1,0)
--         | (x,y)==(-1,0), (0,-1)
--         | (x,y)=(0,-1), (0,-1)
--         | (x,y)=


-- Helpers
turnRight :: (Eq a, Num a)=>(a,a) -> (a,a)
turnRight (1,0)=(0,-1)
turnRight (0,-1)=(-1,0)
turnRight (-1,0)= (0,1)
turnRight (0,1)=(1,0)

turnLeft :: (Eq a, Num a)=>(a,a) -> (a,a)
turnLeft (0,1) = (1,0)
turnLeft (1,0) = (-1,0)
turnleft (-1,0) = (0,-1)
turnleft (0,-1) = (0,1)