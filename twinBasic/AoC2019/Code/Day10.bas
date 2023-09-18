Attribute VB_Name = "Day10"
Option Explicit

Const AN_ASTEROID                           As String = "#"


Public Sub Day10()

    Dim myXYOfAsteroids As Kvp
    'Set myXYOfAsteroids = GetXYOfAsteroids(Part2TestXYMap)
    Set myXYOfAsteroids = GetXYOfAsteroids(GetDay10Input)
    Debug.Print "There are " & myXYOfAsteroids.Count & " asteroids"
    Debug.Print
    
    Dim myBestAsteroid As Asteroid
    Set myBestAsteroid = Day10Part1(myXYOfAsteroids)
    
    Dim myBestAsteroidBearings As Kvp
    Set myBestAsteroidBearings = GetAsteroidBearingsAndDistanceFromOrigin(myBestAsteroid, myXYOfAsteroids)

    Day10Part2 myBestAsteroidBearings
    Debug.Print
    Debug.Print "Finished Day 10"
    
End Sub


Public Function Day10Part1(ByRef ipXYOfAsteroids As Kvp) As Asteroid

    Dim myBestAsteroid As Asteroid
    Dim myMaximumVisibleAsteroids As Long
    myMaximumVisibleAsteroids = 0
    Dim myOriginAsteroid As Asteroid
    Dim myCounter As Long: myCounter = 0
    For Each myOriginAsteroid In ipXYOfAsteroids
        Debug.Print myCounter
        Dim myAsteroidsBearingAndDistancefromOrigin As Kvp
        Set myAsteroidsBearingAndDistancefromOrigin = GetAsteroidBearingsAndDistanceFromOrigin(myOriginAsteroid, ipXYOfAsteroids)
        If myAsteroidsBearingAndDistancefromOrigin.Count > myMaximumVisibleAsteroids Then
        
            Set myBestAsteroid = myOriginAsteroid
            myMaximumVisibleAsteroids = myAsteroidsBearingAndDistancefromOrigin.Count
        
        End If
        myCounter = myCounter + 1
        DoEvents
    Next

    Set Day10Part1 = myBestAsteroid
    Debug.Print "Test x,y should be 11,13", myBestAsteroid.XCoordinate; ","; myBestAsteroid.YCoordinate
    Debug.Print "The Day 10; answer should be 292 ", myMaximumVisibleAsteroids
    
End Function

Public Sub Day10Part2(ByVal ipAsteroidBearings As Kvp)

    PrintAsteroid 0, ipAsteroidBearings
    PrintAsteroid 1, ipAsteroidBearings
    PrintAsteroid 2, ipAsteroidBearings
    PrintAsteroid 3, ipAsteroidBearings
    PrintAsteroid 10, ipAsteroidBearings
    PrintAsteroid 20, ipAsteroidBearings
    PrintAsteroid 50, ipAsteroidBearings
    PrintAsteroid 100, ipAsteroidBearings
    PrintAsteroid 199, ipAsteroidBearings
    PrintAsteroid 200, ipAsteroidBearings
    PrintAsteroid 201, ipAsteroidBearings

    
    Dim myBearings As Variant
    myBearings = ipAsteroidBearings.GetSortedKeys

    Dim myDistanceKvp As Kvp
    Set myDistanceKvp = ipAsteroidBearings.Item(myBearings(199))
    Dim myAsteroid200 As Asteroid
    
    Set myAsteroid200 = myDistanceKvp.GetFirst
    
    Debug.Print "The Day 2 Part 2 answer is 317", myAsteroid200.XCoordinate * 100 + myAsteroid200.YCoordinate


End Sub

Public Sub PrintAsteroid(ByVal ipAsteroid As Long, ByRef ipAsteroidBearings As Kvp)
    Dim myBearings As Variant
    myBearings = ipAsteroidBearings.GetSortedKeys

    Dim myDistancesKvp As Kvp
    Set myDistancesKvp = ipAsteroidBearings.Item(myBearings(ipAsteroid))
    
'    Dim myDistances As Variant
'    myDistances = ipAsteroidBearings.Item(myBearings(ipAsteroid)).GetKeys
'    myDistances = ipAsteroidBearings.Item(myBearings(ipAsteroid)).GetSortedKeys

    Dim myAsteroid As Asteroid
    Set myAsteroid = myDistancesKvp.GetFirst
    
    Debug.Print ipAsteroid, "X,y = "; myAsteroid.XCoordinate & "," & myAsteroid.YCoordinate
    
    
End Sub


Public Function GetAsteroidBearingsAndDistanceFromOrigin(ByRef ipOrigin As Asteroid, ByRef ipXYOfAsteroids As Kvp) As Kvp

    Dim myTargetAsteroid As Asteroid
    Dim myBearings As Kvp: Set myBearings = New Kvp
    For Each myTargetAsteroid In ipXYOfAsteroids
        
        myTargetAsteroid.UpdateBearingAndDistanceFromOrigin ipOrigin
        
        If myTargetAsteroid.Distance > 0 Then
            If myBearings.LacksKey(myTargetAsteroid.Bearing) Then
                myBearings.AddByKey myTargetAsteroid.Bearing, New Kvp
                
            End If
            
            myBearings.Item(myTargetAsteroid.Bearing).AddByKey myTargetAsteroid.Distance, myTargetAsteroid
        End If
        
    Next
    
    Set GetAsteroidBearingsAndDistanceFromOrigin = myBearings
    
End Function


Public Function GetDay10Input() As Kvp

    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Projects\Code\AdventOfCode\Day10\asteroidmap.txt", Scripting.IOMode.ForReading)
        
    Dim myMap  As Kvp: Set myMap = New Kvp
    
    Do
    
        myMap.AddByIndex myfile.ReadLine
        
    Loop Until myfile.AtEndOfStream
        
    myfile.Close
    Set GetDay10Input = myMap
    
End Function

Public Function GetXYOfAsteroids(ByRef ipStringMap As Kvp) As Kvp
        
    Dim myXYMap  As Kvp: Set myXYMap = New Kvp
    Dim myY As Long
    For myY = 0 To ipStringMap.Count - 1

        Dim myX As Long
        For myX = 1 To Len(ipStringMap.GetItem(myY))
            
            If Mid$(ipStringMap.GetItem(myY), myX, 1) = AN_ASTEROID Then
            
                myXYMap.AddByIndex Asteroid.Debut(myX - 1, myY)
                
            End If
            
        Next
        
    Next
        
    Set GetXYOfAsteroids = myXYMap
    
End Function



'Public Sub Day10Tests()
'
'    Dim myMapOfAsteroids As Kvp
'    Set myMapOfAsteroids = GetDay10MapOfAsteroids
'
'    SetupMaths
'    Initialise
'    Dim myIndex As Long
'    For myIndex = 0&To 4&
'
'        s.Map.AddByKey CLngLng(myIndex), New Kvp
'
'    Next
'
'    s.Map.Item(0&).AddByIndexAsChars ".#..#"
'    s.Map.Item(1&).AddByIndexAsChars "....."
'    s.Map.Item(2&).AddByIndexAsChars "#####"
'    s.Map.Item(3&).AddByIndexAsChars "....#"
'    s.Map.Item(4&).AddByIndexAsChars "...##"
'
'
'    FindBestAsteroidSiteBasedOnVisibleAsteroids
'    Debug.Print "Test 1", 8 - s.BestCount
'
'    Set s.Map = New Kvp
'
'    SetupMaths
'    Initialise
'    For myIndex = 0&To 9&
'
'        s.Map.AddByKey myIndex, New Kvp
'
'    Next
'    s.Map.Item(0&).AddByIndexAsChars "......#.#."
'    s.Map.Item(1&).AddByIndexAsChars "#..#.#...."
'    s.Map.Item(2&).AddByIndexAsChars "..#######."
'    s.Map.Item(3&).AddByIndexAsChars ".#.#.###.."
'    s.Map.Item(4&).AddByIndexAsChars ".#..#....."
'    s.Map.Item(5&).AddByIndexAsChars "..#....#.#"
'    s.Map.Item(6&).AddByIndexAsChars "#..#....#."
'    s.Map.Item(7&).AddByIndexAsChars ".##.#..###"
'    s.Map.Item(8&).AddByIndexAsChars "##...#..#."
'    s.Map.Item(9&).AddByIndexAsChars ".#....####"
'
'
'    FindBestAsteroidSiteBasedOnVisibleAsteroids
'    Debug.Print "Test 2", 33 - s.BestCount
'
'
'    Set s.Map = New Kvp
'
'    SetupMaths
'    Initialise
'    For myIndex = 0 To 9
'
'        s.Map.AddByKey myIndex, New Kvp
'
'    Next
'
'    s.Map.Item(0&).AddByIndexAsChars "#.#...#.#."
'    s.Map.Item(1&).AddByIndexAsChars ".###....#."
'    s.Map.Item(2&).AddByIndexAsChars ".#....#..."
'    s.Map.Item(3&).AddByIndexAsChars "##.#.#.#.#"
'    s.Map.Item(4&).AddByIndexAsChars "....#.#.#."
'    s.Map.Item(5&).AddByIndexAsChars ".##..###.#"
'    s.Map.Item(6&).AddByIndexAsChars "..#...##.."
'    s.Map.Item(7&).AddByIndexAsChars "..##....##"
'    s.Map.Item(8&).AddByIndexAsChars "......#..."
'    s.Map.Item(9&).AddByIndexAsChars ".####.###."
'
'
'    FindBestAsteroidSiteBasedOnVisibleAsteroids
'    Debug.Print "Test 3", 35 - s.BestCount
'
'
'Set s.Map = New Kvp
'
'    SetupMaths
'    Initialise
'    For myIndex = 0 To 9
'
'        s.Map.AddByKey myIndex, New Kvp
'
'    Next
'
'    s.Map.Item(0&).AddByIndexAsChars ".#..#..###"
'    s.Map.Item(1&).AddByIndexAsChars "####.###.#"
'    s.Map.Item(2&).AddByIndexAsChars "....###.#."
'    s.Map.Item(3&).AddByIndexAsChars "..###.##.#"
'    s.Map.Item(4&).AddByIndexAsChars "##.##.#.#."
'    s.Map.Item(5&).AddByIndexAsChars "....###..#"
'    s.Map.Item(6&).AddByIndexAsChars "..#.#..#.#"
'    s.Map.Item(7&).AddByIndexAsChars "#..#.#.###"
'    s.Map.Item(8&).AddByIndexAsChars ".##...##.#"
'    s.Map.Item(9&).AddByIndexAsChars ".....#.#.."
'
'
'    FindBestAsteroidSiteBasedOnVisibleAsteroids
'    Debug.Print "Test 4", 41 - s.BestCount
'
'
'
'Set s.Map = New Kvp
'
'    SetupMaths
'    Initialise
'    For myIndex = 0 To 19
'
'        s.Map.AddByKey myIndex, New Kvp
'
'    Next
'
'    s.Map.Item(0&).AddByIndexAsChars ".#..##.###...#######"
'    s.Map.Item(1&).AddByIndexAsChars "##.############..##."
'    s.Map.Item(2&).AddByIndexAsChars ".#.######.########.#"
'    s.Map.Item(3&).AddByIndexAsChars ".###.#######.####.#."
'    s.Map.Item(4&).AddByIndexAsChars "#####.##.#.##.###.##"
'    s.Map.Item(5&).AddByIndexAsChars "..#####..#.#########"
'    s.Map.Item(6&).AddByIndexAsChars "####################"
'    s.Map.Item(7&).AddByIndexAsChars "#.####....###.#.#.##"
'
'    s.Map.Item(8&).AddByIndexAsChars "##.#################"
'    s.Map.Item(9&).AddByIndexAsChars "#####.##.###..####.."
'    s.Map.Item(10&).AddByIndexAsChars "..######..##.#######"
'    s.Map.Item(11&).AddByIndexAsChars "####.##.####...##..#"
'    s.Map.Item(12&).AddByIndexAsChars ".#####..#.######.###"
'    s.Map.Item(13&).AddByIndexAsChars "##...#.##########..."
'    s.Map.Item(14&).AddByIndexAsChars "#.##########.#######"
'    s.Map.Item(15&).AddByIndexAsChars ".####.#.###.###.#.##"
'    s.Map.Item(16&).AddByIndexAsChars "....##.##.###..#####"
'    s.Map.Item(17&).AddByIndexAsChars ".#.#.###########.###"
'    s.Map.Item(18&).AddByIndexAsChars "#.#.#.#####.####.###"
'    s.Map.Item(19&).AddByIndexAsChars "###.##.####.##.#..##"
'
'    FindBestAsteroidSiteBasedOnVisibleAsteroids
'    Debug.Print "Test 5", 210 - s.BestCount
'
'End Sub

















Public Function Part2TestXYMap() As Kvp

    Dim myStringMap As Kvp: Set myStringMap = New Kvp
    
    myStringMap.AddByKey 0&, ".#..##.###...#######"
    myStringMap.AddByKey 1&, "##.############..##."
    myStringMap.AddByKey 2&, ".#.######.########.#"
    myStringMap.AddByKey 3&, ".###.#######.####.#."
    myStringMap.AddByKey 4&, "#####.##.#.##.###.##"
    myStringMap.AddByKey 5&, "..#####..#.#########"
    myStringMap.AddByKey 6&, "####################"
    myStringMap.AddByKey 7&, "#.####....###.#.#.##"
    myStringMap.AddByKey 8&, "##.#################"
    myStringMap.AddByKey 9&, "#####.##.###..####.."
    myStringMap.AddByKey 10&, "..######..##.#######"
    myStringMap.AddByKey 11&, "####.##.####...##..#"
    myStringMap.AddByKey 12&, ".#####..#.######.###"
    myStringMap.AddByKey 13&, "##...#.##########..."
    myStringMap.AddByKey 14&, "#.##########.#######"
    myStringMap.AddByKey 15&, ".####.#.###.###.#.##"
    myStringMap.AddByKey 16&, "....##.##.###..#####"
    myStringMap.AddByKey 17&, ".#.#.###########.###"
    myStringMap.AddByKey 18&, "#.#.#.#####.####.###"
    myStringMap.AddByKey 19&, "###.##.####.##.#..##"
    
    Set Part2TestXYMap = myStringMap

End Function
