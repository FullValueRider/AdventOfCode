Attribute VB_Name = "Day08"
Option Explicit

Const ImageSize As Long = 25 * 6


Public Sub Day8Part1And2()

    Dim myImagelayers As Collection
    Set myImagelayers = GetImageData(ImageSize)
    
    Dim myMin0Layer As Long
    myMin0Layer = GetMin0Layer(myImagelayers)
        
    'Part1
    Debug.Print "The Day 8 part 1 answer should be 1620", GetNumberCount(1, myImagelayers.Item(myMin0Layer)) * GetNumberCount(2, myImagelayers.Item(myMin0Layer))
    
    'Part2
    Debug.Print " The Day 8 Part 2 Day 2 answer should be BCYEF": Debug.Print: Debug.Print:
    PrintImage GetMessageLayer(myImagelayers)
    Debug.Print: Debug.Print "Day 8 Complete"
End Sub


'@Ignore AssignedByValParameter
Public Sub PrintImage(ByVal ipMessage As String)
    
    ipMessage = Replace(ipMessage, "0", " ")
    Do While Len(ipMessage) > 0
    
        Debug.Print VBA.Left$(ipMessage, 26)
        ipMessage = VBA.Mid$(ipMessage, 26)
        
    Loop
    
End Sub


Public Function GetMessageLayer(ByRef ipLayers As Collection) As String

    Dim myMessageLayer As String
    Dim myChar As Long
    For myChar = 1 To Len(ipLayers.Item(1))
    
        Dim myLayer As Variant
        For Each myLayer In ipLayers
        
            If Mid$(myLayer, myChar, 1) <> "2" Then
            
                myMessageLayer = myMessageLayer & Mid$(myLayer, myChar, 1)
                Exit For
                
            End If
        
        Next
        
    Next
    
    GetMessageLayer = myMessageLayer
        
End Function


Public Function GetMin0Layer(ByRef ipLayers As Collection) As Long

    Dim myZeros As Long: myZeros = 0
    Dim myMinZeroIndex As Long: myMinZeroIndex = 0
    Dim myMinZeros As Long: myMinZeros = -1
    Dim myItem As Long
    For myItem = 1 To ipLayers.Count
    
        Dim myLayer As String
        myLayer = ipLayers.Item(myItem)
        
        If myMinZeros = -1 Then myMinZeros = Len(myLayer) + 1
        myZeros = GetNumberCount(0, myLayer)
        
        If myZeros < myMinZeros Then
        
            myMinZeroIndex = myItem
            myMinZeros = myZeros
            
        End If
        
    Next
    
    GetMin0Layer = myMinZeroIndex

End Function


Public Function GetNumberCount(ByVal ipNumber As Long, ByVal ipLayer As String) As Long

        GetNumberCount = Len(ipLayer) - Len(Replace(ipLayer, CStr(ipNumber), vbNullString))

End Function


Public Function GetImageData(ByVal ipLayerSize As Long) As Collection

    Dim myFso  As Scripting.FileSystemObject
    Set myFso = New Scripting.FileSystemObject
    
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile("C:\Users\slayc\source\repos\VBA\AdventOfCode\2019\Day08Input.txt", Scripting.IOMode.ForReading)
        
    Dim myLayers  As Collection: Set myLayers = New Collection
    Dim myLayer As String
    Do
        
        myLayer = myfile.Read(ipLayerSize)
        
        If Len(myLayer) < ipLayerSize Then
        
            Debug.Print "Image data is not a multiple of layer size)"
            End
            
        End If
        
        myLayers.Add myLayer
    
    Loop Until myfile.AtEndOfStream
    
    myfile.Close
    Set GetImageData = myLayers
    
End Function

