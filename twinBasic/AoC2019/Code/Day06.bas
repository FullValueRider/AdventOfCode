Attribute VB_Name = "Day06"
Option Explicit

'@Folder("AdventOfCode")

Private Const Origin                                    As String = "COM"

Private Type State

    Orbits                                              As Kvp
    Bodies                                              As Kvp
    Source                                              As Kvp
    
End Type

Private s                                           As State

Public Sub Day6CodePart1()
    
    SetupOrbits
    Debug.Print GetTotalOrbits
    
End Sub


Public Sub Day6Part2()

    SetupOrbits
    
    Dim SantaOrbit As Kvp
    Set SantaOrbit = GetOrbitalPath("SAN")
    
    Dim myOrbit As Kvp
    Set myOrbit = GetOrbitalPath("YOU")
    
    Dim myIntersect As Kvp
    Set myIntersect = SantaOrbit.Cohorts(myOrbit)
    
    Dim myMovement As Long
    myMovement = SantaOrbit.Count + myOrbit.Count - (2 * myIntersect.Item(CohortType_KeysInAandB).Count) - 2
    Debug.Print myMovement
    
End Sub

Public Sub SetupOrbits()

    Initialise
    
    Do While s.Source.Count > 0
    
        Dim myItem As Variant
        
        For Each myItem In s.Source.GetKeys
            
            If s.Bodies.HoldsKey(myItem) Then
            
                AddBody s.Source.Item(myItem)
                s.Source.Remove myItem
                Exit For
                
            End If
        
        Next
    
    Loop

End Sub


Public Function GetOrbitalPath(ByVal ipName As String) As Kvp

    Dim myPath As Kvp
    Set myPath = New Kvp
    myPath.AddByKey ipName, s.Bodies.Item(ipName)
    
    Do
        
        Dim myLastKey As String
        myLastKey = myPath.GetKeys(myPath.Count - 1)
        myPath.AddByKey myPath.Item(myLastKey).Parent.Name, myPath.Item(myLastKey).Parent
        
    Loop Until myPath.HoldsKey(Origin)
        
    Set GetOrbitalPath = myPath
    
End Function


Public Sub Initialise()

    Set s.Source = GetInputKvp
    
    Dim myBody  As Day6Orbit
    Set myBody = New Day6Orbit

    With myBody

        Set .Children = s.Orbits
        Set .Parent = myBody
        .Position = 0
        .Name = Origin

    End With
    
    Set s.Orbits = New Kvp
    s.Orbits.AddByKey Origin, myBody
    
    Set s.Bodies = New Kvp
    s.Bodies.AddByKey Origin, myBody
    AddBody s.Source.Item("COM")
    s.Source.Remove "COM"
   
End Sub


Public Sub AddBody(ByVal ipBody As String)

    Dim myHost As String
    myHost = Split(ipBody, ")")(0)
    
    Dim myPlanets As Variant
    myPlanets = Split(Split(ipBody, ")")(1), ",")
    
    Dim myPlanet  As Variant
    For Each myPlanet In myPlanets
    
        Dim myOrbit As Day6Orbit
        Set myOrbit = New Day6Orbit
        Set myOrbit.Parent = s.Bodies.Item(myHost)
        If myOrbit.Children Is Nothing Then Set myOrbit.Children = New Kvp
        ' the code belkow is wrong but we are getting the correct answer
        If myOrbit.Children.Count = 0 Then
            
            myOrbit.Position = myOrbit.Parent.Children.Item(myHost).Position + 1
            
        Else
        
            myOrbit.Position = myOrbit.Children.Item(myOrbit.Children.GetKeys(myOrbit.Children.Count - 1)).Position + 1
        
        End If
        
        myOrbit.Name = myPlanet
        myOrbit.Children.AddByKey myPlanet, myOrbit
        s.Bodies.AddByKey myPlanet, myOrbit.Children.Item(myPlanet)
 
    Next

End Sub


'@Ignore FunctionReturnValueNotUsed
Public Function GetTotalOrbits() As Long

    Dim myBody As Variant
    For Each myBody In s.Bodies
    
        Dim myOrbit As Day6Orbit
        Set myOrbit = myBody
        
        Dim myCount As Long
        myCount = myCount + myOrbit.Position
        
    Next
    
    GetTotalOrbits = myCount
    
End Function

Public Function GetInputKvp() As Kvp

    Dim myInput As String
    '@Ignore AssignmentNotUsed
    myInput = "97W)B43,R63)RTM,19C)SHD,F85)8Z8,5D9)Z5T,RG5)R48,HJC)NVP,46W)1SL,BFP)34P,2M5)NWQ,CJD)M9C,8BK)QK5,TWT)53F,KT1)7FJ,WG8)TBN,FMZ)2RY,"
    myInput = myInput & "734)H1Y,8XM)GSZ,TXP)Z9D,NLX)TNW,7ST)N94,W68)1WS,S1J)TDP,5C2)8C4,ZW4)GLV,MHN)CVN,4FH)NRP,4SC)155,6D9)PG7,L44)CZ6,KC5)PX6,PFD)MGG,"
    myInput = myInput & "5JJ)GBH,63X)QYN,N96)VM4,WGN)JMD,Y9P)RS9,N6C)3HD,4KF)1JP,M31)3RJ,W71)9C5,FLV)M31,J6Q)HSP,ZPF)27T,YXX)MJZ,Q43)F38,FPM)599,TJX)M6S,"
    myInput = myInput & "N94)5L3,6LN)T2D,9QK)S49,N81)PD5,VG4)264,MS6)HYV,JBP)HCJ,SV7)C2F,LZT)NCW,7FJ)SV7,TYJ)S5S,YJY)2WW,KKL)JDY,QBL)5VX,R7W)C7X,1KP)JLP,"
    myInput = myInput & "NKQ)9C6,BVR)LWN,3ZW)F31,Q86)188,V5D)CX3,7L7)RKS,6HC)LCZ,T2F)8HT,KFX)H6G,D1Z)NXM,5NH)PFD,FRR)TWD,GVY)KJW,GBX)LVT,KGZ)CNH,GPZ)7J8,"
    myInput = myInput & "4QY)6LN,9VD)4QY,4LY)SH6,P9Y)RJW,WK5)1HQ,SPS)667,VDN)1VF,HCJ)ZRW,SLV)C4Y,MS6)76G,LPD)Y9P,CZ6)2C1,VGT)J3F,YS4)FMZ,SNR)KFL,Z4P)4FJ,"
    myInput = myInput & "WQW)WBX,13Z)NXK,3Z2)VX9,XZJ)KDN,BH9)HNB,7W1)F78,B3H)H2Q,ZN3)G75,3VY)NN9,2W3)9F7,X8L)NKQ,1J7)41S,DN3)NV7,NY1)KSJ,HFG)NXQ,27T)QHD,"
    myInput = myInput & "VX9)4WC,TTJ)MQT,NVH)JJ7,9L5)D3H,VJM)MJD,XZ3)N81,83R)ZSZ,9S5)FYR,CMZ)W1P,2L2)XRP,XP9)JSF,JM9)H51,3GG)3M8,6D9)L44,5G2)2RN,1KK)ZS3,"
    myInput = myInput & "68S)JMG,TDP)3W9,XXY)LMB,8GD)JVN,H9L)7N1,7J8)KQJ,HNB)RN8,SGV)DS4,B2P)2M5,HSP)XXY,4JJ)SXX,1ZZ)HYW,MVQ)F27,RMW)P9Y,HLR)HF7,6MS)X73,"
    myInput = myInput & "5Z5)VQY,XS9)VDN,3VT)38P,LPB)BQH,1RT)XS9,9F7)K3K,G75)3H5,6CD)B2F,BTY)2PK,2BH)1WY,Y8Y)1QJ,B9G)3LD,172)65H,8HC)YDG,PX6)41H,X4P)FH4,"
    myInput = myInput & "JSF)6SW,ZBY)XGM,2M1)SYH,W6B)3VY,MGG)48J,R41)C9C,B43)V78,M84)9S5,PTF)23B,ZJZ)33X,ZBY)YHL,56M)21K,7F8)3R6,GY6)QBL,CM3)P9B,LWN)JZ8,"
    myInput = myInput & "NWQ)NK4,9MB)M8X,D12)TWT,8WF)KKL,RB9)4C4,23B)93Z,ZQZ)FSP,QMW)DN3,HRL)GDF,1PN)795,B7L)DYL,CF8)68K,FYR)JP2,X4D)DB5,NM2)39M,FKQ)G4J,"
    myInput = myInput & "Y9D)91R,SHD)56S,C5X)9KN,24S)4JJ,R48)KG1,8NS)QD3,72X)WD9,WWH)Q86,1MF)99H,RMB)QY1,BNP)Y6V,FJS)13R,KXS)QH8,GM7)YC5,VMM)B3T,M8X)3GG,"
    myInput = myInput & "815)H7Q,KSJ)CM3,VLP)T8B,JDY)X4P,PQH)CNL,1QX)FTN,CTV)VB6,HYW)LGL,HLC)7VN,L54)YJ8,NKP)ZJZ,4GQ)1CL,96B)PDN,S37)7L4,2RY)BKW,D5P)XF3,"
    myInput = myInput & "MF3)S26,8DN)NNN,7L4)5H7,CD4)H4M,H5D)JM7,C7D)T2F,SMF)W61,J7P)F22,41H)8CF,DQR)C8H,Q3D)417,GRX)MVQ,5GN)BZB,9SY)K2Y,X57)XH6,JQX)B8Q,"
    myInput = myInput & "78L)85W,PDN)13Z,21K)SRY,D5V)ZYC,VXF)PF3,815)184,76G)X44,FDW)B9G,KJW)7JQ,667)YQW,F31)QLK,JNK)W7V,46W)HCB,56S)F3Y,93Z)1T7,RBQ)KRY,"
    myInput = myInput & "894)XXW,33X)XG5,591)KCF,X44)JQX,J91)7L7,DDZ)MKC,7W2)5TX,M6S)7J9,H7G)CK7,145)2VG,FJS)MCC,PJT)NK1,7BY)MC3,GBH)33J,F4Y)FVB,JFX)K8R,"
    myInput = myInput & "5VX)815,VWG)PLJ,JHH)QJ4,4G6)R2K,KL7)J5Y,XXM)5NH,H5L)Y4Z,85W)8XV,YFK)PTV,4R3)8P5,TVW)MZK,V6C)KKX,P23)DPN,8V4)SC9,SH6)JCF,PTF)1DP,"
    myInput = myInput & "JCF)XHT,1VR)ZKZ,FFW)RPW,QN1)MBD,F7K)V92,VM4)ZD2,FTN)2MC,L5V)ZR9,COM)1MF,XZ3)6H3,KGP)4FH,M77)R62,2R2)GTL,SYH)JXP,98K)F4P,6K3)GNT,"
    myInput = myInput & "BWB)Q3D,VV8)YOU,MC3)SF5,T1V)1CQ,37V)TJX,CNL)PJT,9YW)FX6,862)NZW,J7B)PVT,WXN)X57,SXX)6DP,ZKZ)BWB,XNW)9XX,ZQZ)7W1,F27)NZV,C7X)5D9,"
    myInput = myInput & "64K)JNP,5H7)HXC,XDJ)GXK,33J)TBK,1JP)7WX,LHQ)TQK,YJC)2LJ,4YL)72X,GXK)BTF,1CL)X87,VF9)KC5,8CF)KQ7,2P8)3VT,Z11)D1R,2PK)1YT,DLM)T93,"
    myInput = myInput & "2VG)NM2,R96)CC3,YHL)J7B,K41)ZN3,G6N)L3V,57Y)NLJ,MG8)P8P,2C5)5GN,3G3)SQ3,WBX)XT7,SRY)R63,79R)VG9,BTF)N1J,LYR)W68,6QN)4TH,YRT)BS8,"
    myInput = myInput & "MM8)Y37,BCW)BDZ,MCC)CH2,44Y)RMB,PF3)V7K,ZR9)HSN,YRZ)Z19,5TG)SL6,NDL)PQH,63K)73H,V3C)KXS,JLM)14R,4TH)4KF,J76)5N1,WJN)NKP,TV6)19C,"
    myInput = myInput & "Z8F)4S7,F6W)1VR,417)QFZ,67J)GSQ,TYV)GDL,JMD)4BY,7PX)FYX,K8R)W9L,LGL)QMW,KPX)V7D,53F)FJS,3GH)61Z,QYD)JNK,VSP)NWK,JNP)2M1,VLH)FDW,"
    myInput = myInput & "KVT)85X,NSK)FFW,D2J)NVK,B54)Z82,5MP)1R5,J38)SMF,2QG)3GY,37V)CDG,XPN)6XR,HCB)MPY,ZX9)WK5,5WF)SRB,DLS)G7D,FGS)5MP,1T7)X6X,KG1)9LV,"
    myInput = myInput & "QTB)F37,9LV)JH9,3RJ)8DJ,FRB)MM8,KKX)LZT,HH9)RF7,D1R)89F,QJ4)1L6,DT9)KGP,N4S)P1P,F52)S37,XH6)Y27,TXV)YFJ,ND8)ZX9,43N)F4Y,NNY)S1J,"
    myInput = myInput & "PG7)BNP,RPW)WWH,4S3)VJB,Z82)JGP,BQY)64L,SQ3)91C,D3V)XGG,3Z2)FPM,TW8)Y8L,HSN)S4N,H6G)DMP,SHS)SWG,5N7)6QN,M8X)WNH,6V1)PPM,1R5)MF3,"
    myInput = myInput & "H9B)CFF,GSZ)WGN,LZQ)GPZ,RP6)2C5,X87)D3V,PTV)P9X,G1T)YRZ,2NW)C42,BQH)JXN,CFJ)2L2,MVW)LQQ,SRB)JM9,YXL)HD2,7J9)6V1,ZPF)B3H,GSJ)NBY,"
    myInput = myInput & "C8H)Y8Y,1BS)H5L,XGG)6MS,KZF)734,DD1)8G9,QHL)95X,3LD)9H1,D8Q)1KK,XWH)Y4M,P8P)2BH,188)C2G,8XV)CFJ,V7K)H7G,4G5)W2M,TPX)9WY,MJZ)V2W,"
    myInput = myInput & "YC5)F1T,W61)H25,9P3)8R8,1XZ)642,DXD)TXP,MJZ)SLV,CN8)NKZ,LYS)ZNL,MQT)J38,LKZ)TSQ,HNQ)9XZ,226)84R,8HT)591,4BY)6X1,VLH)XTX,"
    myInput = myInput & "DPD)RG5,QLK)HZL,92P)CTV,GMF)78T,2RY)RKV,9NQ)HRL,YDG)YX4,5X3)V7S,8G9)QP9,GDF)6HC,5TX)FD4,4QT)MSV,13R)NPK,GNT)ZBY,4S7)MCY,"
    myInput = myInput & "N4S)DK9,VYW)8PD,ZSN)24S,JGP)J91,YS8)YFK,XXW)QN1,8R7)92P,FD4)T1V,XHT)PZB,DYV)VSP,JJ7)N8D,WPM)NLR,QHD)GBK,5R9)1V4,QYN)PXJ,"
    myInput = myInput & "TZM)XXM,172)WY2,NB5)GGC,ZRW)9VV,F5K)WR3,FSS)P4Z,T2D)4G5,1X9)YRT,C7V)F46,NLJ)HNQ,3DV)1GD,1NB)G1T,1ML)HH9,YZK)JBP,TDP)8BK,"
    myInput = myInput & "R62)H9L,NWK)98Y,J55)7ST,SF5)BFP,MQD)D5L,XT7)P1Y,JHB)MG9,SSM)8K8,NXQ)W5P,MBT)B3D,P1P)YJY,C9C)BCW,GQ6)KFX,QP9)HM4,PG7)XNW,"
    myInput = myInput & "XXW)ZHS,V84)LPD,XG5)9CV,QRL)V5D,8R8)DT9,FBK)G25,1SK)37V,L8Q)38M,GCX)QW1,JP9)ZW4,W15)GH9,9VV)WWN,MDP)97C,NCW)JFR,GBD)9PM,"
    myInput = myInput & "Q77)7PX,ZB1)XPN,88Y)B7L,N96)862,TBK)5R9,R2K)T1W,DPN)BS7,96X)4YL,Z8F)4R3,S4N)TZM,PLJ)78L,GGC)T65,BZ2)H5D,YJY)L8Q,34P)1ML,"
    myInput = myInput & "HXC)LZQ,GXF)SFV,FYR)6B4,T93)FFK,39M)9NQ,1V4)NTH,6PG)W15,TV6)C2K,MYH)HYC,1PN)VLP,62K)JNV,PG8)ZL9,RHY)DPD,N1J)MQN,C36)894,"
    myInput = myInput & "264)ZB1,CH2)FPT,R7R)QG8,4WC)VG4,22N)Y5Y,NSQ)8V4,NTH)M77,L6J)GBD,QML)4QT,F82)V6C,Y55)2J3,FYX)P23,LJW)5QL,3XW)GBW,JRN)RB9,"
    myInput = myInput & "MQN)563,48J)293,3CP)JHB,VV7)5TG,1GD)8LL,D8H)7BY,61Z)4GX,F78)KL7,L3V)NSQ,TKH)2QG,F85)YS4,XPN)43N,4RF)W7C,YFJ)226,92P)SAN,"
    myInput = myInput & "N9J)G6N,HF7)9LR,H68)6D9,6H3)HLC,887)1PN,2C1)XT3,CS2)SSM,CNH)4GQ,QD3)635,G7D)172,NLR)VV7,3CZ)N6C,GTL)J6Q,D3H)D99,19C)VGT,"
    myInput = myInput & "T75)FGS,MGY)96X,J5Y)BKR,3RJ)VDX,RKV)376,FFK)XP9,F96)B2P,G4J)FVT,QFZ)1JH,3L3)RHY,3PN)2G8,YJ8)5N7,14R)X7P,H1T)SHG,1XB)RMZ,"
    myInput = myInput & "ZVB)H68,ZHS)CJD,DW1)9MB,795)5P9,84R)2R2,KDN)QTB,38P)JP9,2WW)2FM,C4Y)67J,5V3)4M3,LMB)6N4,DYL)ND8,98Y)BQY,ZYC)JRN,1FX)R41,"
    myInput = myInput & "9CV)PGG,S5S)68S,PMB)1G6,ZKX)9S2,Z18)1XB,JH9)ZL5,12K)WV2,FVB)9VD,5PN)XWM,JP2)4RF,XRP)3Z2,F96)YS8,B3D)5C2,M6S)BH9,F86)145,"
    myInput = myInput & "WRP)C5X,MBD)5PN,WV9)RMW,T9S)VXF,JVN)WG8,SC9)VYW,91C)BM1,Y3Y)RPN,91R)X3J,NN3)CMZ,HYC)LPB,WH4)WRP,6B4)1SK,BT7)HJ8,MBT)4LY,"
    myInput = myInput & "LQQ)8DN,BDZ)F9M,SYH)MGY,XF3)JFJ,1BV)79R,J6Q)QBZ,3HD)TYJ,7WX)P34,226)MBT,1HQ)DLS,SSM)NY1,PZB)8NS,NKZ)YJC,W7V)Y3Y,Y6V)TVW,"
    myInput = myInput & "QBZ)QRL,8DJ)GCX,FZC)3GH,4R1)ZKX,WNH)B54,FYJ)L5V,887)F96,7TV)JDD,W68)17B,VDX)FBK,BW8)MQD,54Y)3L3,GLV)1J7,353)MYG,D2J)T9S,"
    myInput = myInput & "8K8)2P8,3H5)B6N,1YT)RBQ,V78)1RT,YQW)SHJ,RJW)VF9,GRH)TYV,2FM)TW8,X3J)LHQ,B8Q)D6G,293)3CP,3GY)R7R,PBS)MHN,1JH)WH4,XTX)4J9,"
    myInput = myInput & "B3J)8VM,KSJ)3XW,1LR)F55,3W9)4YX,Y6V)LYR,7MF)1XZ,F46)BVG,635)DQR,P4Z)5HT,MCL)1BV,JXN)G22,6CD)1BS,JMG)9SY,CDG)7W2,VB6)DLM,"
    myInput = myInput & "JRN)7TN,599)BMS,XT6)MVW,9WY)9YW,N94)42C,R9Z)NNY,YQW)HFN,MJD)Y55,2M1)VD6,TBN)2JQ,68G)YQN,C2F)HQG,ZNL)GXF,7TN)DPT,557)9P3,"
    myInput = myInput & "S49)FKQ,L6Y)VJM,C2S)MCL,SH1)MG8,G25)JLM,TSQ)HLR,FHC)D2J,RPN)F85,SFV)XT6,NZV)6PG,QH8)JFX,NNN)GRH,642)KT1,N8D)8Y9,S66)F82,"
    myInput = myInput & "V92)YTT,B3T)T75,MVW)YHZ,F1T)NDL,PVT)1ZZ,56F)NN3,JFJ)7MF,S2K)LJW,P9Y)98K,F3Y)5X3,1G6)N9J,9XX)SNR,XDD)YXX,68K)MYH,3R6)63X,"
    myInput = myInput & "L3J)KBZ,MFG)XHN,2G8)DDZ,GH9)4N5,BKW)B27,KDD)1FX,6QN)PG8,K3K)46W,HM4)1Y5,1BS)XZ3,RG5)WXT,G8F)HYT,M9C)XDJ,Y8L)WJN,MKC)9L5,"
    myInput = myInput & "1VF)F6W,RXR)97W,ZCW)9ZR,RZF)G9J,JMD)1KP,X73)NSK,FX6)MXL,ZL5)3DV,RN8)LYS,RF7)56F,376)GVY,8C4)K41,H51)5Z5,4M3)64K,JGP)J55,"
    myInput = myInput & "5QL)L54,7N1)FYJ,XWN)MWH,Q3D)4R1,3M8)54Y,P9B)5KD,44Y)95Q,ZL9)XWH,SGK)4G6,3NJ)1QX,Y5Y)557,TWD)WV9,KZF)Q1Q,QY1)16J,NK1)C2S,"
    myInput = myInput & "HQG)VMM,836)2W3,345)44Y,97C)SSB,X8L)JG9,TQK)F5K,8DN)9FD,5LT)Z11,5P9)6K3,WY2)56M,P23)CD4,YJC)MS6,5KD)57Y,7J9)WPM,1DP)8V1,"
    myInput = myInput & "JNV)NLX,M5P)BW8,YHZ)G8F,VJB)X8L,6X1)L3J,W2M)FN4,C2K)JHH,JLP)C7D,GBW)887,184)H1T,LVT)LGV,SHG)TV6,9XG)FHC,BRC)CF8,KFL)353,"
    myInput = myInput & "NZW)FRB,FN2)5WF,KBZ)LKZ,5N1)3ZW,8PD)22N,Q1Q)CN8,2F4)FN2,B2F)PMB,1SL)1X9,Y4M)Z4P,YQN)T8N,JV3)4S3,9C5)DD1,B84)5V3,2RN)3L6,"
    myInput = myInput & "XTF)KCX,Y4Z)TTJ,T8B)W6B,YX4)8XM,JFR)5JJ,HQ4)JVD,9PM)2NW,9C6)QHL,P1Y)12K,BZB)VV8,1WY)62K,DPT)VLH,NPK)5LT,RPN)7F8,9S2)V84,"
    myInput = myInput & "WR3)BZ2,89F)X4D,5HT)V6N,HYV)PBS,T65)88Y,F22)D5P,HZL)XWN,JDD)M84,8P5)Q43,5G3)3G3,B6N)RXR,WXT)BVR,2JQ)68G,95X)SHS,JBP)B55,"
    myInput = myInput & "4YX)D5V,KCF)Y7L,ZWF)126,F37)D8Q,JZ8)6CD,ZDK)D12,2MC)GY6,6XR)QSY,WV2)FRR,G22)SGK,KCX)SH1,155)C7V,183)HFG,K41)BT7,BH9)ZCW,"
    myInput = myInput & "1Y5)S66,1L6)RZF,P34)DW1,Y9Q)B3J,42C)7FR,VQY)3PN,NWK)J76,D99)JV3,914)ZVB,NN9)YXL,GRH)WXN,DK9)D4H,DS4)Y4H,D5L)KZF,C2G)ZSN,"
    myInput = myInput & "7VN)F7X,9NQ)3NJ,64L)FZC,S26)Y9Q,4N5)CS2,X7P)5G2,RKS)SGV,PNK)W71,9KN)R7W,K2Y)9R5,8J4)63K,9H1)TPX,LMB)HY1,NVK)N96,5ZC)B95,"
    myInput = myInput & "PPM)5G3,CNL)KGZ,JG9)D5J,17B)Y9D,XHN)VWG,8Z8)HF4,32R)D8H,F38)XX1,PLJ)TKH,38M)MDP,HYT)HQ4,NBY)WQW,GSQ)GBX,XT3)2F4,V7S)D1Z,"
    myInput = myInput & "3ZW)3CZ,WWN)F7K,NVP)1LR,KRY)GSJ,G9J)D4W,JXP)PFQ,Z9D)5ZC,LYR)KPX,FSP)8HC,16J)Q77,MPY)1NB,F78)R9Z,RN8)H9B,F3N)XTF,PD5)96B,"
    myInput = myInput & "5L3)183,X97)RP6,VD6)ZWF,F4P)GRX,M77)836,4FJ)M3N,HFN)8J4,NRP)GM7,QSY)C36,MSV)XZJ,126)L6Y,99H)XDD,Z5T)BTY,TNW)8WF,HJ8)4SC,"
    myInput = myInput & "9LR)BRC,F7X)SPS,1VR)B84,9R5)GMF,H1Y)NB5,6DP)TXV,4J9)N4S,8Y2)PTF,YBJ)ZDK,1WS)83R,BMS)FSS,MG9)FLV,BM1)QML,1SK)KDD,SF5)X5V,"
    myInput = myInput & "3HD)9QK,P9X)ZR1,YC5)HJC,BKR)V3C,WD9)ZPF,TNW)S2K,GDL)Z18,5D9)F52,ZD2)NVH,F55)345,6N4)7TV,X5V)9XG,64K)YBJ,8Y9)ZQZ,LCZ)MFG,"
    myInput = myInput & "H4M)Z8F,V2W)J7P,73H)H9J,V7D)M5P,FN4)PNK,NK4)R96,T1W)914,HY1)HWY,H5D)8Y2,3NJ)QYD,2LJ)DYV,SL6)KVT,H2Q)DXD,CFF)L6J,D5J)F86,"
    myInput = myInput & "41S)8R7,H9J)YZK,QW1)X97,B27)GMP,8V1)GQ6,8VM)F3N,Y4H)8GD,SHJ)32R"
    
    
    If InStr(myInput, Origin) = 0 Then
    
        Debug.Print "There is no origin for the orbit descriptions"
        End

    End If

    Dim myArray As Variant
    myArray = Split(myInput, ",")
    
    Dim myKvp As Kvp
    Set myKvp = New Kvp
    
    Dim myItem As Variant
    For Each myItem In myArray
    
        If Len(myItem) <> 7 Then
            
            Debug.Print "Bad orbit description " & myItem
            End
            
        End If
        
        Dim myPair As Variant
        myPair = Split(myItem, ")")
        
        If myKvp.HoldsKey(myPair(0)) Then
        
            Dim mySystem    As String
            mySystem = myKvp.Item(myPair(0)) & "," & myPair(1)
            myKvp.Remove myPair(0)
            myKvp.AddByKey myPair(0), mySystem
            
        Else
        
            myKvp.AddByKey myPair(0), myItem
            
        End If
        
    Next
    
    Set GetInputKvp = myKvp
    
End Function
