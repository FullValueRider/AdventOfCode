VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@PredeclaredId
Private Const RawData As String = "bgvyzdsv"

Public Sub Part1()

    Dim myLowestNumber As Long
    myLowestNumber = GetLowestNumberStartingWith("00000")
    Debug.Print "The answer to Day 4 Part 1 is 254575. Found is ", myLowestNumber
    
End Sub


Public Sub Part2()

    Dim myLowestNumber As Long
    myLowestNumber = GetLowestNumberStartingWith("000000")
    Debug.Print "The answer to Day 4 Part 2 is XXXXXX. Found is ", myLowestNumber
    
End Sub


Public Function GetLowestNumberStartingWith(ByVal myZeros As String) As Long

    Dim myNum As Long
    myNum = -1
    Dim myLen  As Long
    myLen = Len(myZeros)
    
    Do
    
        DoEvents
        myNum = myNum + 1
        
        Dim myHash As String
        myHash = StringToMD5Hex(RawData & CStr(myNum))
        
    Loop Until VBA.Left$(myHash, myLen) = myZeros
        
    GetLowestNumberStartingWith = myNum
    
End Function


Public Function StringToMD5Hex(ByVal s As String) As String
    ' Nicked from https://www.mrexcel.com/board/threads/convert-string-to-md5-hash.973381/

    Dim enc As Object
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    Dim myBytes() As Byte
    myBytes = VBA.StrConv(s, vbFromUnicode)
    myBytes = enc.ComputeHash_2(myBytes)

    Dim myPos As Long
    Dim myOutstr As String
    For myPos = LBound(myBytes) To UBound(myBytes)
        myOutstr = myOutstr & VBA.LCase$(VBA.Right$("0" & VBA.Hex$(myBytes(myPos)), 2))
    Next

    StringToMD5Hex = myOutstr
    Set enc = Nothing
    
End Function
