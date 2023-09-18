Attribute VB_Name = "Day14"
'@IgnoreModule
Option Explicit
'    Dim OreEaters As KvpOD: Set OreEaters = New KvpOD
'    Dim Reaction As KvpOD: Set Resctions = New KvpOD
Const OreEater                              As Long = 0&
Const Route                                 As Long = 1&

Const IsOre                                 As String = "ORE"


Public Function ConsolidateToOreConsumers(ByVal ipSorted As KvpOD) As KvpOD


    Dim myOreEaters As KvpOD: Set myOreEaters = New KvpOD
    
    Dim myItem As Variant
    
    For Each myItem In ipSorted.Item(OreEater)
    
        Dim myEater As OreConsumer: Set myEater = New OreConsumer
       ' Set myEater = OreEater.Debutante(myItem)
        myOreEaters.AddByKey myEater.Name, myEater
        
    Next
    
    Set ConsolidateToOreConsumers = myOreEaters

End Function


Public Function TestInput() As KvpOD

    Dim myKvp As KvpOD: Set myKvp = New KvpOD
    
    myKvp.AddByIndex "157 ORE => 5 NZVS"
    myKvp.AddByIndex "165 ORE => 6 DCFZ"
    myKvp.AddByIndex "44 XJWVT, 5 KHKGT, 1 QDVJ, 29 NZVS, 9 GPVTF, 48 HKGWZ => 1 FUEL"
    myKvp.AddByIndex "12 HKGWZ, 1 GPVTF, 8 PSHF => 9 QDVJ"
    myKvp.AddByIndex "179 ORE => 7 PSHF"
    myKvp.AddByIndex "177 ORE => 5 HKGWZ"
    myKvp.AddByIndex "7 DCFZ, 7 PSHF => 2 XJWVT"
    myKvp.AddByIndex "165 ORE => 2 GPVTF"
    myKvp.AddByIndex "3 DCFZ, 7 NZVS, 5 HKGWZ, 10 PSHF => 8 KHKGT"
    
    Set TestInput = myKvp
End Function


Public Function GetSortedInput(ByVal myStrings As KvpOD) As KvpOD

    Dim mySorted As KvpOD: Set mySorted = New KvpOD
    mySorted.AddByKey OreEater, New KvpOD
    mySorted.AddByKey Route, New KvpOD
    
    Dim myItem As Variant
    For Each myItem In myStrings
    
        If InStr(myItem, IsOre) = 0 Then
        
            mySorted.Item(Route).AddByIndex myItem
        
        Else
        
            mySorted.Item(OreEater).AddByIndex myItem
        
        End If
        
    Next
    
    Set GetSortedInput = mySorted
    
End Function
