Attribute VB_Name = "Day1"
Option Explicit
'@Folder("AdventOfCode")


Public Sub PrepForLaunch()

   Debug.Print ComponentFuel(1969)
    Debug.Print CalculateFuel
End Sub

'@Ignore FunctionReturnValueNotUsed
Public Function CalculateFuel() As Long


Dim ComponentMasses             As Variant
Dim myFuel                      As Long
Dim myComponent                 As Variant


    
    ComponentMasses = _
        Array(54296, 106942, 137389, 116551, 129293, 60967, 142448, 101720, 64463, 142264, 68673, _
        144661, 110426, 59099, 63711, 120365, 125233, 126793, 61990, 122059, 86768, 134293, 114985, _
        61280, 75325, 103102, 116332, 112075, 114895, 98816, 59389, 124402, 74995, 135512, 115619, _
        59680, 61421, 141160, 148880, 70010, 119379, 92155, 126698, 138653, 149004, 142730, 68658, _
        73811, 87064, 62684, 93335, 140475, 143377, 98445, 117960, 80237, 132483, 108319, 104154, _
        99383, 104685, 114888, 73376, 58590, 132759, 114399, 77796, 119228, 136282, 84789, 66511, _
        51939, 142313, 117305, 139543, 92054, 64606, 139795, 109051, 97040, 91850, 107391, 60200, _
        75812, 74898, 64884, 115210, 85055, 92256, 67470, 90286, 129142, 109235, 117194, 104028, _
        127482, 68502, 92440, 50369, 84878)

    For Each myComponent In ComponentMasses
    
        myFuel = myFuel + ComponentFuel(CLng((myComponent) \ 3) - 2)
    
    
    Next
    
    CalculateFuel = myFuel
    
End Function


Public Function ComponentFuel(ByVal ipComponent As Long) As Long

    If ipComponent <= 0 Then
    
        ComponentFuel = 0
        Exit Function
        
    Else
        
        ComponentFuel = ipComponent + ComponentFuel((ipComponent \ 3) - 2)
        
    End If
        
End Function


'Public Sub TestCOmputer()
'
'Dim myProgram As Variant
''@Ignore VariableNotUsed, VariableNotAssigned
'Dim myItem          As Variant
'Dim noun As Long
'Dim verb As Long
'
'    For noun = 0 To 99
'
'    For verb = 0 To 99
'        myProgram = ComputeProgram(noun, verb, Array(1, 0, 0, 3, 1, 1, 2, 3, 1, 3, 4, 3, 1, 5, 0, 3, 2, 6, 1, 19, 1, 19, 10, 23, 2, 13, 23, 27, 1, 5, 27, 31, 2, 6, 31, 35, 1, 6, 35, 39, 2, 39, 9, 43, 1, 5, 43, 47, 1, 13, 47, 51, 1, 10, 51, 55, 2, 55, 10, 59, 2, 10, 59, 63, 1, 9, 63, 67, 2, 67, 13, 71, 1, 71, 6, 75, 2, 6, 75, 79, 1, 5, 79, 83, 2, 83, 9, 87, 1, 6, 87, 91, 2, 91, 6, 95, 1, 95, 6, 99, 2, 99, 13, 103, 1, 6, 103, 107, 1, 2, 107, 111, 1, 111, 9, 0, 99, 2, 14, 0, 0))
'        If myProgram(0) = 19690720 Then
'
'            Debug.Print 100 * noun + verb
'            Exit Sub
'        End If
'    Next
'
'    Next
'
'End Sub





Public Sub ManhattenDistance()

Dim Wire1   As Variant
Dim Wire2   As Variant
Dim Wire1Path As String
Dim Wire2path As String


    '@Ignore AssignmentNotUsed
    Wire1Path = "R999 , D666, L86, U464, R755, U652, R883, D287, L244, U308, L965, U629, R813, U985, R620, D153, L655, D110, R163, D81, L909," _
        & "D108 , L673, D165, L620, U901, R601, D561, L490, D21, R223, U478, R80, U379, R873, U61, L674, D732, R270, U297, L354, U264, L615," _
        & "D2, R51, D582, R280, U173, R624, U644, R451, D97, R209, U245, R32, U185, R948, D947, R380, D945, L720, U305, R911, U614, L419, D751," _
        & "L934, U371, R291, D166, L137, D958, R368, U441, R720, U822, R961, D32, R242, D972, L782, D166, L680, U111, R379, D155, R213, U573," _
        & "R761, D543, R762, U953, R317, U841, L38, U900, R573, U766, R807, U950, R945, D705, R572, D994, L633, U33, L173, U482, R253, D835, R800," _
        & "U201, L167, U97, R375, D813, L468, D924, L972, U570, R975, D898, L195, U757, L565, D378, R935, U4, L334, D707, R958, U742, R507, U892," _
        & "R174, D565, L862, D311, L770, D619, L319, D698, L169, D652, L761, D644, R837, U43, L197, D11, L282, D345, L551, U460, R90, D388, R911," _
        & "U602, L21, D275, L763, U880, R604, D838, R146, U993, L99, U99, R928, U54, L148, D863, R618, U449, R97,L135,D966,R121,U763,R46,D110,R830," _
        & "U644,L932,D122,L123,U145,R273,U690,L443,D372,R818,D259,L695,U69,R73,D718,R106,U929,L346,D291,L857,D341,R297,D823,R819,U496,L958,U394," _
        & "R102,D763,L444,D835,L33,U45,R812,U845,R196,U458,R231,U637,R661,D983,L941,U975,L353,U609,L698,U152,R122,D882,R682,D926,R729,U429,R255," _
        & "D227,R987,D547,L446,U217,R678,D464,R849,D472,L406,U940,L271,D779,R980,D751,L171,D420,L49,D271,R430,D530,R509,U479,R135,D770,R85,U815,R328,U234,R83"
    
    Wire1Path = Replace(Wire1Path, " ", vbNullString)
    Wire1 = Split(Wire1Path, ",")
        
    '@Ignore AssignmentNotUsed
    Wire2path = "L1008,D951,L618,U727,L638,D21,R804,D19,L246,U356,L51,U8,L627,U229,R719,D198,L342,U240,L738,D393,L529,D22,R648,D716,L485,U972,L580," _
        & "U884,R612,D211,L695,U731,R883,U470,R732,U723,R545,D944,R18,U554,L874,D112,R782,D418,R638,D296,L123,U426,L479,U746,L209,D328,L121,D496," _
        & "L172,D228,L703,D389,R919,U976,R364,D468,L234,U318,R912,U236,R148,U21,R26,D116,L269,D913,L949,D206,L348,U496,R208,U706,R450,U472,R637,U884," _
        & "L8,U82,L77,D737,L677,D358,L351,U719,R154,U339,L506,U76,L952,D791,L64,U879,R332,D244,R638,D453,L107,D908,L58,D188,R440,D147,R913,U298,L681," _
        & "D582,L943,U503,L6,U459,L289,D131,L739,D443,R333,D138,R553,D73,L475,U930,L332,U518,R614,D553,L515,U602,R342,U95,R131,D98,R351,U921,L141,U207," _
        & "R199,U765,R55,U623,R768,D620,L722,U31,L891,D862,R85,U271,R590,D184,R960,U149,L985,U82,R591,D384,R942,D670,R584,D637,L548,U844,R353,U496,L504," _
        & "U3,L830,U239,R246,U279,L146,U965,R784,U448,R60,D903,R490,D831,L537,U109,R271,U306,L342,D99,L234,D936,R621,U870,R56,D29,R366,D562,R276,D134," _
        & "L289,D425,R597,D102,R276,D600,R1,U322,L526,D744,L259,D111,R994,D581 , L973, D871, R173, D924, R294, U478, R384, D242, R606, U629, R472, D651," _
        & "R526, U55, R885, U637, R186, U299, R812, D95, R390, D689, R514, U483, R471, D591, L610, D955, L599, D674, R766, U834, L417, U625, R903, U376," _
        & "R991 , U175, R477, U524, L453, D407, R72, D217, L968, D892, L806, D589, R603, U938, L942, D940, R578, U820, L888, U232, L740, D348, R445, U269," _
        & "L170, U979, L159, U433, L31, D818, L914, U600, L33, U159, R974, D983, L922, U807, R682, U525, L234, U624, L973, U123, L875, D64, L579, U885," _
        & "L911, D578, R17, D293, L211"
    
    Wire2path = Replace(Wire2path, " ", vbNullString)
    Wire2 = Split(Wire2path, ",")
    
Dim Wire1Coords As Kvp
Dim Wire2Coords As Kvp

    Set Wire1Coords = GetXY(Wire1)
    Set Wire2Coords = GetXY(Wire2)
    Debug.Print Wire1Coords.Count
    Debug.Print Wire2Coords.Count
    
Dim myItem                      As Variant
Dim myCrossXY                   As Kvp

    Set myCrossXY = New Kvp
    
    For Each myItem In Wire2Coords.GetKeys
        
        If Wire1Coords.HoldsKey(myItem) Then
            'Debug.Print myItem
            If myCrossXY.LacksKey(myItem) Then
            
                myCrossXY.AddByKey myItem, Wire1Coords.Item(myItem) + Wire2Coords.Item(myItem)
                Debug.Print myItem, Wire1Coords.Item(myItem), Wire2Coords.Item(myItem), myCrossXY.Item(myItem)
                
            End If
            
        End If
        
    Next
    
'   Debug.Print myCrossXY.Count
'
'Dim myManhatten                     As Long
'Dim mySmallest                      As Long
'Dim distX                           As Long
'Dim distY                           As Long
'
'
'    mySmallest = 0
'
'    For Each myItem In myCrossXY
'
'        distX = Abs(CLng(Split(myItem, ",")(0)))
'        distY = Abs(CLng(Split(myItem, ",")(1)))
'        myManhatten = distX + distY
'        Debug.Print myManhatten
'
'    Next
   
End Sub


Public Function GetXY(ByVal ipPathArray As Variant) As Kvp

Dim myCoord                         As Variant
Dim myX                             As Long
Dim myY                             As Long
Dim myXY                            As Kvp
Dim myPosition                      As String
Dim myStep                          As Long

    myX = 0
    myY = 0
    Set myXY = New Kvp
    myStep = 0
    
    For Each myCoord In ipPathArray
    
        Dim myDirection As String
        Dim myCount As Long
        
        myDirection = Left$(myCoord, 1)
        myCount = CLng(Mid$(myCoord, 2))
        
        
        '@Ignore VariableNotUsed, UndeclaredVariable
        For myStep = 1 To myCount
            
            Select Case myDirection
            
                Case "U": myY = myY + 1
                Case "D": myY = myY - 1
                Case "R": myX = myX + 1
                Case "L": myX = myX - 1
                Case Else
                
                    Debug.Print "Unknown direction " & myDirection
            
            End Select
            
            myStep = myStep + 1
            myPosition = CStr(myX) & "," & CStr(myY)
            
            If myXY.LacksKey(myPosition) Then
                
                myXY.AddByKey myPosition, myStep
            
            End If
            
        Next
            
    Next
    
    Set GetXY = myXY
    
End Function









