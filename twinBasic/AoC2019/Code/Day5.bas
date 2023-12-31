Attribute VB_Name = "Day5"
Option Explicit

'@Folder("AdventOfCode")

'@Ignore ConstantNotUsed
Const Temp As String = "Temp"


'Private Sub Tester()
'
'    'Using position AccessMode, consider whether the input is equal to 8; output 1 (if it is) or 0 (if it is not).
'    s.Code = Array(3, 9, 8, 9, 10, 9, 4, 9, 99, -1, 8)
'    s.Input = Array(8)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 1", 1, s.Output
'
'    s.Code = Array(3, 9, 8, 9, 10, 9, 4, 9, 99, -1, 8)
'    s.Input = Array(-1)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 2", 0, s.Output
'
'    'Using position AccessMode, consider whether the input is less than 8; output 1 (if it is) or 0 (if it is not).
'    s.Code = Array(3, 9, 7, 9, 10, 9, 4, 9, 99, -1, 8)
'    s.Input = Array(0)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 3", 1, s.Output
'
'
'    s.Code = Array(3, 9, 7, 9, 10, 9, 4, 9, 99, -1, 8)
'    s.Input = Array(8)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 4", 0, s.Output
'
'    'Using immediate AccessMode, consider whether the input is equal to 8; output 1 (if it is) or 0 (if it is not).
'    s.Code = Array(3, 3, 1108, -1, 8, 3, 4, 3, 99)
'    s.Input = Array(7)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 5", 0, s.Output
'
'    s.Code = Array(3, 3, 1108, -1, 8, 3, 4, 3, 99)
'    s.Input = Array(8)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 6", 1, s.Output
'
'    ' Using immediate AccessMode, consider whether the input is less than 8; output 1 (if it is) or 0 (if it is not).
'    s.Code = Array(3, 3, 1107, -1, 8, 3, 4, 3, 99)
'    s.Input = Array(-1)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 7", 1, s.Output
'
'    s.Code = Array(3, 3, 1107, -1, 8, 3, 4, 3, 99)
'    s.Input = Array(9)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 8", 0, s.Output
'
'    'using position AccessModeoutput 0 if the input was zero or 1 if the input was non-zero:
'    s.Code = Array(3, 12, 6, 12, 15, 1, 13, 14, 13, 4, 13, 99, -1, 0, 1, 9)
'    s.Input = Array(0)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 9", 0, s.Output
'
'    s.Code = Array(3, 12, 6, 12, 15, 1, 13, 14, 13, 4, 13, 99, -1, 0, 1, 9)
'    s.Input = Array(9)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 10", 1, s.Output
'
'
'    'using immediate AccessMode 0 if the input was zero or 1 if the input was non-zero:
'    s.Code = Array(3, 3, 1105, -1, 9, 1101, 0, 0, 12, 4, 12, 99, 1)
'    s.Input = Array(0)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 11", 0, s.Output
'
'    s.Code = Array(3, 3, 1105, -1, 9, 1101, 0, 0, 12, 4, 12, 99, 1)
'    s.Input = Array(9)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 12", 1, s.Output
'
'
'    'output 999 if the input value is below 8, output 1000 if the input value is equal to 8, or output 1001 if the input value is greater than 8
'    s.Code = Array(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99)
'    s.Input = Array(0)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 13", 999, s.Output
'
'    s.Code = Array(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99)
'    s.Input = Array(8)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 14", 1000, s.Output
'
'    s.Code = Array(3, 21, 1008, 21, 8, 20, 1005, 20, 22, 107, 8, 21, 20, 1006, 20, 31, 1106, 0, 36, 98, 0, 0, 1002, 21, 125, 20, 4, 20, 1105, 1, 46, 104, 999, 1105, 1, 46, 1101, 1000, 1, 20, 4, 20, 1105, 1, 46, 98, 99)
'    s.Input = Array(1080)
'    s.Output = vbNullString
'    Run
'    Debug.Print "Test 15", 1001, s.Output
'
'End Sub


'Private Sub Part1()
'
'    s.Code = Array(3, 225, 1, 225, 6, 6, 1100, 1, 238, 225, 104, 0, 1101, 90, 64, 225, 1101, 15, 56, 225, 1, 14, 153, 224, 101, -147, 224, 224, 4, 224, _
'            1002, 223, 8, 223, 1001, 224, 3, 224, 1, 224, 223, 223, 2, 162, 188, 224, 101, -2014, 224, 224, 4, 224, 1002, 223, 8, 223, 101, 6, 224, _
'            224, 1, 223, 224, 223, 1001, 18, 81, 224, 1001, 224, -137, 224, 4, 224, 1002, 223, 8, 223, 1001, 224, 3, 224, 1, 223, 224, 223, 1102, _
'            16, 16, 224, 101, -256, 224, 224, 4, 224, 1002, 223, 8, 223, 1001, 224, 6, 224, 1, 223, 224, 223, 101, 48, 217, 224, 1001, 224, -125, _
'            224, 4, 224, 1002, 223, 8, 223, 1001, 224, 3, 224, 1, 224, 223, 223, 1002, 158, 22, 224, 1001, 224, -1540, 224, 4, 224, 1002, 223, _
'            8, 223, 101, 2, 224, 224, 1, 223, 224, 223, 1101, 83, 31, 225, 1101, 56, 70, 225, 1101, 13, 38, 225, 102, 36, 192, 224, 1001, 224, _
'            -3312, 224, 4, 224, 1002, 223, 8, 223, 1001, 224, 4, 224, 1, 224, 223, 223, 1102, 75, 53, 225, 1101, 14, 92, 225, 1101, 7, 66, 224, _
'            101, -73, 224, 224, 4, 224, 102, 8, 223, 223, 101, 3, 224, 224, 1, 224, 223, 223, 1101, 77, 60, 225, 4, 223, 99, 0, 0, 0, 677, 0, 0, 0, _
'            0, 0, 0, 0, 0, 0, 0, 0, 1105, 0, 99999, 1105, 227, 247, 1105, 1, 99999, 1005, 227, 99999, 1005, 0, 256, 1105, 1, 99999, 1106, 227, _
'            99999, 1106, 0, 265, 1105, 1, 99999, 1006, 0, 99999, 1006, 227, 274, 1105, 1, 99999, 1105, 1, 280, 1105, 1, 99999, 1, 225, 225, 225, _
'            1101, 294, 0, 0, 105, 1, 0, 1105, 1, 99999, 1106, 0, 300, 1105, 1, 99999, 1, 225, 225, 225, 1101, 314, 0, 0, 106, 0, 0, 1105, 1, 99999, _
'            7, 226, 677, 224, 1002, 223, 2, 223, 1005, 224, 329, 1001, 223, 1, 223, 1007, 226, 677, 224, 1002, 223, 2, 223, 1005, 224, 344, 101, _
'            1, 223, 223, 108, 226, 226, 224, 1002, 223, 2, 223, 1006, 224, 359, 101, 1, 223, 223, 7, 226, 226, 224, 102, 2, 223, 223, 1005, 224, _
'            374, 101, 1, 223, 223, 8, 677, 677, 224, 1002, 223, 2, 223, 1005, 224, 389, 1001, 223, 1, 223, 107, 677, 677, 224, 102, 2, 223, 223, _
'            1006, 224, 404, 101, 1, 223, 223, 1107, 677, 226, 224, 102, 2, 223, 223, 1006, 224, 419, 1001, 223, 1, 223, 1008, 226, 226, 224, 1002 _
'            , 223, 2, 223, 1005, 224, 434, 1001, 223, 1, 223, 7, 677, 226, 224, 102, 2, 223, 223, 1006, 224, 449, 1001, 223, 1, 223, 1107, 226, 226, _
'            224, 1002, 223, 2, 223, 1005, 224, 464, 101, 1, 223, 223, 1108, 226, 677, 224, 102, 2, 223, 223, 1005, 224, 479, 101, 1, 223, 223, 1007, _
'            677, 677, 224, 102, 2, 223, 223, 1006, 224, 494, 1001, 223, 1, 223, 1107, 226, 677, 224, 1002, 223, 2, 223, 1005, 224, 509, 101, 1, 223, _
'            223, 1007, 226, 226, 224, 1002, 223, 2, 223, 1006, 224, 524, 101, 1, 223, 223, 107, 226, 226, 224, 1002, 223, 2, 223, 1005, 224, 539, _
'            1001, 223, 1, 223, 1108, 677, 677, 224, 1002, 223, 2, 223, 1005, 224, 554, 101, 1, 223, 223, 1008, 677, 226, 224, 102, 2, 223, 223, _
'            1006, 224, 569, 1001, 223, 1, 223, 8, 226, 677, 224, 102, 2, 223, 223, 1005, 224, 584, 1001, 223, 1, 223, 1008, 677, 677, 224, 1002, _
'            223, 2, 223, 1006, 224, 599, 1001, 223, 1, 223, 108, 677, 677, 224, 102, 2, 223, 223, 1006, 224, 614, 1001, 223, 1, 223, 108, 226, _
'            677, 224, 102, 2, 223, 223, 1005, 224, 629, 101, 1, 223, 223, 8, 677, 226, 224, 102, 2, 223, 223, 1005, 224, 644, 101, 1, 223, 223, _
'            107, 677, 226, 224, 1002, 223, 2, 223, 1005, 224, 659, 101, 1, 223, 223, 1108, 677, 226, 224, 102, 2, 223, 223, 1005, 224, 674, 1001, 223, 1, 223, 4, 223, 99, 226)
'    's.Code = Array(1002, 4, 3, 4, 33)
'    s.Input = Array(1)
'    s.Output = vbNullString
'    Run
'    Debug.Print "7988899", s.Output
'
'End Sub
'
'Private Sub Part2()
'
'    s.Code = Array(3, 225, 1, 225, 6, 6, 1100, 1, 238, 225, 104, 0, 1101, 90, 64, 225, 1101, 15, 56, 225, 1, 14, 153, 224, 101, -147, 224, 224, 4, 224, _
'            1002, 223, 8, 223, 1001, 224, 3, 224, 1, 224, 223, 223, 2, 162, 188, 224, 101, -2014, 224, 224, 4, 224, 1002, 223, 8, 223, 101, 6, 224, _
'            224, 1, 223, 224, 223, 1001, 18, 81, 224, 1001, 224, -137, 224, 4, 224, 1002, 223, 8, 223, 1001, 224, 3, 224, 1, 223, 224, 223, 1102, _
'            16, 16, 224, 101, -256, 224, 224, 4, 224, 1002, 223, 8, 223, 1001, 224, 6, 224, 1, 223, 224, 223, 101, 48, 217, 224, 1001, 224, -125, _
'            224, 4, 224, 1002, 223, 8, 223, 1001, 224, 3, 224, 1, 224, 223, 223, 1002, 158, 22, 224, 1001, 224, -1540, 224, 4, 224, 1002, 223, _
'            8, 223, 101, 2, 224, 224, 1, 223, 224, 223, 1101, 83, 31, 225, 1101, 56, 70, 225, 1101, 13, 38, 225, 102, 36, 192, 224, 1001, 224, _
'            -3312, 224, 4, 224, 1002, 223, 8, 223, 1001, 224, 4, 224, 1, 224, 223, 223, 1102, 75, 53, 225, 1101, 14, 92, 225, 1101, 7, 66, 224, _
'            101, -73, 224, 224, 4, 224, 102, 8, 223, 223, 101, 3, 224, 224, 1, 224, 223, 223, 1101, 77, 60, 225, 4, 223, 99, 0, 0, 0, 677, 0, 0, 0, _
'            0, 0, 0, 0, 0, 0, 0, 0, 1105, 0, 99999, 1105, 227, 247, 1105, 1, 99999, 1005, 227, 99999, 1005, 0, 256, 1105, 1, 99999, 1106, 227, _
'            99999, 1106, 0, 265, 1105, 1, 99999, 1006, 0, 99999, 1006, 227, 274, 1105, 1, 99999, 1105, 1, 280, 1105, 1, 99999, 1, 225, 225, 225, _
'            1101, 294, 0, 0, 105, 1, 0, 1105, 1, 99999, 1106, 0, 300, 1105, 1, 99999, 1, 225, 225, 225, 1101, 314, 0, 0, 106, 0, 0, 1105, 1, 99999, _
'            7, 226, 677, 224, 1002, 223, 2, 223, 1005, 224, 329, 1001, 223, 1, 223, 1007, 226, 677, 224, 1002, 223, 2, 223, 1005, 224, 344, 101, _
'            1, 223, 223, 108, 226, 226, 224, 1002, 223, 2, 223, 1006, 224, 359, 101, 1, 223, 223, 7, 226, 226, 224, 102, 2, 223, 223, 1005, 224, _
'            374, 101, 1, 223, 223, 8, 677, 677, 224, 1002, 223, 2, 223, 1005, 224, 389, 1001, 223, 1, 223, 107, 677, 677, 224, 102, 2, 223, 223, _
'            1006, 224, 404, 101, 1, 223, 223, 1107, 677, 226, 224, 102, 2, 223, 223, 1006, 224, 419, 1001, 223, 1, 223, 1008, 226, 226, 224, 1002 _
'            , 223, 2, 223, 1005, 224, 434, 1001, 223, 1, 223, 7, 677, 226, 224, 102, 2, 223, 223, 1006, 224, 449, 1001, 223, 1, 223, 1107, 226, 226, _
'            224, 1002, 223, 2, 223, 1005, 224, 464, 101, 1, 223, 223, 1108, 226, 677, 224, 102, 2, 223, 223, 1005, 224, 479, 101, 1, 223, 223, 1007, _
'            677, 677, 224, 102, 2, 223, 223, 1006, 224, 494, 1001, 223, 1, 223, 1107, 226, 677, 224, 1002, 223, 2, 223, 1005, 224, 509, 101, 1, 223, _
'            223, 1007, 226, 226, 224, 1002, 223, 2, 223, 1006, 224, 524, 101, 1, 223, 223, 107, 226, 226, 224, 1002, 223, 2, 223, 1005, 224, 539, _
'            1001, 223, 1, 223, 1108, 677, 677, 224, 1002, 223, 2, 223, 1005, 224, 554, 101, 1, 223, 223, 1008, 677, 226, 224, 102, 2, 223, 223, _
'            1006, 224, 569, 1001, 223, 1, 223, 8, 226, 677, 224, 102, 2, 223, 223, 1005, 224, 584, 1001, 223, 1, 223, 1008, 677, 677, 224, 1002, _
'            223, 2, 223, 1006, 224, 599, 1001, 223, 1, 223, 108, 677, 677, 224, 102, 2, 223, 223, 1006, 224, 614, 1001, 223, 1, 223, 108, 226, _
'            677, 224, 102, 2, 223, 223, 1005, 224, 629, 101, 1, 223, 223, 8, 677, 226, 224, 102, 2, 223, 223, 1005, 224, 644, 101, 1, 223, 223, _
'            107, 677, 226, 224, 1002, 223, 2, 223, 1005, 224, 659, 101, 1, 223, 223, 1108, 677, 226, 224, 102, 2, 223, 223, 1005, 224, 674, 1001, 223, 1, 223, 4, 223, 99, 226)
'    s.Input = Array(5)
'    s.Output = vbNullString
'    Run
'    Debug.Print s.Output
'
'End Sub






