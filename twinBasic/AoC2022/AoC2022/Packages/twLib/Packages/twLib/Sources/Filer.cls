VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Filer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Helpers")
Option Explicit
'@PredeclaredId
'@Exposed
Const ForReading                            As Integer = 1% 'Scripting.IOMode.ForReading

Public Function GetFileAsArrayOfStrings(ByVal ipFilePath As String, Optional ByRef ipSplitAtToken As String = vbCrLf) As Variant

    Dim myFso  As Object
    Set myFso = New Scripting.FileSystemObject
        
    Dim myfile As Object
    Set myfile = myFso.OpenTextFile(ipFilePath, ForReading)
            
    GetFileAsArrayOfStrings = VBA.Split(myfile.ReadAll, ipSplitAtToken)
        
    myfile.Close
        
End Function

Public Function GetFileAsString(ByVal ipPath As String) As String

    Dim myFso  As Object
    Set myFso = New Scripting.FileSystemObject
    Dim myfile As Object
    Set myfile = myFso.OpenTextFile(ipPath, ForReading)
        
        
    GetFileAsString = myfile.ReadAll

    myfile.Close
        
End Function
