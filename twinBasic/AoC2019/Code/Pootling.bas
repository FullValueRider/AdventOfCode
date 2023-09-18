Attribute VB_Name = "Pootling"
Option Explicit

Public Sub AssignToKvpItem()

    Dim myKvp As Kvp: Set myKvp = New Kvp

    myKvp.AddByIndex 42&
    myKvp.AddByIndex 84&

    Debug.Print myKvp.Item(CLng(1))
    Debug.Print myKvp.Item(CLng(2))

    myKvp.Item(1&) = 99

    Debug.Print myKvp.Item(1&)

    TestPassedKvpAssignment myKvp

End Sub

Public Sub TestPassedKvpAssignment(ByVal ipKvp As Kvp)

    Debug.Print ipKvp.Item(1&)
    'This line produces an error 424 'object required
    ipKvp.Item(1&) = 1001
    ' So 1001 is not assigned
    Debug.Print ipKvp.Item(1&)

    Dim myObj As Object: Set myObj = ipKvp
    'This line works fine
    myObj.Item(1&) = 2002
    Debug.Print ipKvp.Item(1&)

End Sub

