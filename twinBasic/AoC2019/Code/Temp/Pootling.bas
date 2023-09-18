Attribute VB_Name = "Pootling"
Option Explicit
'@IgnoreModule
'@Ignore MoveFieldCloserToUsage, EncapsulatePublicField
Public myProvider As EventProvider

Public Sub KvpTestIng()

    Debug.Print "KvpOD test"
    
    Dim myODTest As KvpOD
    Set myODTest = New KvpOD
    
    myODTest.AddByIndexFromArray Split("Hello1 there1 world1 its1 a1 nice1 day1")
    
    ' KvoOD check locals to see if vars are in KvpEnumTest
    Dim myODKeys As Variant
    myODKeys = myODTest.GetKeys
    Debug.Print myODTest.GetKeysAsString
    
    
    Dim myODValues As Variant
    myODValues = myODTest.GetValues
    Debug.Print myODTest.GetValuesAsString
    
    'Test assignemt through .Item
    myODTest.Item(2&) = "Success"
    Debug.Print myODTest.Item(2&)
    
    Dim myODItem As Variant
    For Each myODItem In myODTest
        
        Debug.Print myODItem.Key, myODItem.Value
         
    Next
    
    '
    Debug.Print "Testing wrapped KvpOD"
    Dim myWrapTest As KvpEnumTest
    Set myWrapTest = New KvpEnumTest
    
    myWrapTest.AddByIndexFromArray Split("Hello2 there2 world2 its2 a2 nice2 day2")
    
    ' check locals to see if vars are in KvpEnumTest
    Dim myWrapKeys As Variant
    myWrapKeys = myWrapTest.Keys
    Debug.Print myWrapTest.GetKeysAsString
    
    
    Dim myWrapValues As Variant
    myWrapValues = myWrapTest.Values
    Debug.Print myWrapTest.GetValuesAsString
    
    'Test assignemt through .Item
    myWrapTest.Item(2&) = "Success"
    Debug.Print myWrapTest.Item(2&)
    '
    
    Dim myWrapItem As Variant
    For Each myWrapItem In myWrapTest
        
        Debug.Print myWrapItem.Key, myWrapItem.Value
         
    Next
    
End Sub


Public Sub testKvpForeach()
    'On Error Resume Next
    Dim myInt As Integer: myInt = 1
    Dim myLong As Long: myLong = 4
    Dim myBool As Boolean: myBool = True
    Dim myDouble As Double: myDouble = 3.142
    
    Dim myKvp As KvpOD: Set myKvp = New KvpOD
    Dim myObj As Object
    Set myObj = myKvp
    Debug.Print Err.Number
    Debug.Print Err.Description
    myKvp.AddByIndexFromArray Array(myInt, myLong, myBool, myDouble)
    Debug.Print myKvp.Count
    Debug.Print myKvp.GetKeysAsString(",")
    Debug.Print myKvp.GetValuesAsString(",")
    'Test item method
    
    Dim myResult As Variant
    myResult = myKvp.Item(2&)
    
    myObj.Item(2&) = 42
    myKvp.Item(2&) = 42
    myResult = myKvp.Item(2&)
    'Dim myObj1 As Object
    KvpPassedTest myKvp
    Dim myItem As Variant
    Dim myPair As KVPair
    
    For Each myItem In myKvp
        Set myPair = myItem
        'Debug.Print Err.Number
        'Debug.Print Err.Description
        Debug.Print myPair.Key, myPair.Value
    Next
    
    'Debug.Print myKvp.Item(2^)
    Dim myTest As Long
    myTest = 10
    'myKvp.Item( 2^, myTest
    'Debug.Print "SHould be 10", myKvp.Item(2^)
    myKvp.Remove 2^
    'myKvp.AddByKey 2^, 10
    For Each myItem In myKvp
        Debug.Print myItem.Key
    Next
    
    Dim myKvp2 As KvpOD
    Set myKvp2 = New KvpOD
    myKvp2.AddByKey "one", 1
    myKvp2.AddByKey "two", 4
    myKvp2.AddByKey "three", 8
    
    For Each myItem In myKvp2
        Debug.Print myItem
    Next
End Sub

Public Sub KvpPassedTest(ByVal ipKvp As KvpOD)
    Const MY_TEST As Long = 2
    Debug.Print ipKvp.Item(2&)
    Debug.Print ipKvp.Item(MY_TEST)
End Sub
    

