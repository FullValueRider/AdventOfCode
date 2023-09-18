Attribute VB_Name = "Common"
Option Explicit

Private Type State

    Pi                          As Double
    Tau                         As Double '2 * pi
    RadiansToDegrees            As Double
    
End Type

Private s                       As State


Public Function MakeLongLongVsLongLongKvp(ByVal ipArray As Variant) As Kvp
' TheIntComputer  program manipluates 64 bit numbers so we need to make
' sure that both keys and items are longlong values
' Hence the use of addbykey rather than addbyindex
    
    
    Dim myItem As Variant
    Dim myKvp As Kvp: Set myKvp = New Kvp
    Dim myIndex As LongLong
    myIndex = 0^
    For Each myItem In ipArray
    
        myKvp.AddByKey myIndex, CLngLng(myItem)
        myIndex = myIndex + 1^
        
    Next
    
    Set MakeLongLongVsLongLongKvp = myKvp
    
End Function


Public Function Pi() As Double

    If s.Pi = 0 Then s.Pi = 4 * Atn(1)
    Pi = s.Pi
    
End Function


Public Function Tau() As Double

    If s.Tau = 0 Then s.Tau = 8 * Atn(1)
    Tau = s.Tau
    
End Function


Public Function RadiansToDegrees() As Double

    If s.RadiansToDegrees = 0 Then s.RadiansToDegrees = 180 / (4 * Atn(1))
    RadiansToDegrees = s.RadiansToDegrees
    
End Function






Public Function DistanceFromOrigin(ByVal ipOriginX As Long, ByVal ipOriginY As Long, ByVal ipCurrentX As Long, ByVal ipCurrentY As Long) As Single
    
    If (ipOriginX = ipCurrentX) And (ipOriginY = ipCurrentY) Then
    
        DistanceFromOrigin = 0#
        Exit Function
        
    Else
    
        DistanceFromOrigin = CSng(Sqr(CDbl(ipCurrentX - ipOriginX) ^ 2 + CDbl(ipCurrentY - ipOriginY) ^ 2))
    
    End If
    
End Function


Public Function BearingFromOrigin(ByVal ipOriginX As Long, ByVal ipOriginY As Long, ByVal ipCurrentX As Long, ByVal ipCurrentY As Long) As Single

    Dim myReturn As Single
    myReturn = CSng(CheckForNESW(ipOriginX, ipOriginY, ipCurrentX, ipCurrentY))
    ' myReturn produces a result between 0 and 360 for legal values
    ' -2 means that the current locatoopn is the same as the origina.
    If myReturn = -2 Then

        myReturn = CSng(myReturn)
        
    ElseIf myReturn = 0 Or myReturn = 90 Or myReturn = 180 Or myReturn = 270 Then
    
        myReturn = CSng(myReturn)
        
    Else
    
        Dim ipDeltaX As LongLong
        ipDeltaX = (ipCurrentX - ipOriginX)

        Dim ipDeltaY As LongLong
        ipDeltaY = (ipCurrentY - ipOriginY) '* -1
        
        Dim myTanmyReturn As Single
        myTanmyReturn = Atn(ipDeltaY / ipDeltaX) * RadiansToDegrees
        
        'Debug.Print ipDeltaX; ipDeltaY; myTanmyReturn;
        
        If ipDeltaX > 0 And ipDeltaY > 0 Then  '++
        
            myReturn = 90 + myTanmyReturn  'mytanmyReturn is 0 to 90
            
        ElseIf ipDeltaX > 0 And ipDeltaY < 0 Then '+-
        
            myReturn = 90 + myTanmyReturn  ''mytanmyReturn is -90 to 0
        
        ElseIf ipDeltaX < 0 And ipDeltaY < 0 Then '--
        
             myReturn = 270 + myTanmyReturn  'mytanmyReturn is 0 to 90
        
        Else '+-
        
            myReturn = 270 + myTanmyReturn  'mytanmyReturn is -90 to 0
        
        End If
        
    End If
    
    
'    If myreturn < 180 Then myreturn = Abs(myreturn - 360)
'
'    If myreturn > 180 Then myreturn = 180 - Abs(myreturn - 360)
    'myreturn = myreturn + 270
    If myReturn >= 360 Then myReturn = myReturn - 360
    
    BearingFromOrigin = myReturn
    
    
End Function


Public Sub testxy()
    
'    Debug.Print "Beading should be -2 ", BearingFromOrigin(0, 0, 0, 0)
'    Debug.Print
    
    Debug.Print "Bearing is ", BearingFromOrigin(0, 5, 0, 0)
    Debug.Print
    Debug.Print "Bearing is ", , BearingFromOrigin(2, 5, 0, 0)
    Debug.Print "Bearing is  ", BearingFromOrigin(5, 5, 0, 0)
    Debug.Print "Bearing is ", , BearingFromOrigin(5, 2, 0, 0)
    Debug.Print
    Debug.Print "Bearing is ", BearingFromOrigin(5, 0, 0, 0)
    Debug.Print
    Debug.Print "Bearing is ", , BearingFromOrigin(5, -2, 0, 0)
    Debug.Print "Bearing is  ", BearingFromOrigin(5, -5, 0, 0)
    Debug.Print "Bearing is ", , BearingFromOrigin(2, -5, 0, 0)
    Debug.Print
    Debug.Print "Bearing is ", BearingFromOrigin(0, -5, 0, 0)
    Debug.Print
    Debug.Print "Bearing is  ", , BearingFromOrigin(-2, -5, 0, 0)
    Debug.Print "Bearing is  ", BearingFromOrigin(-5, -5, 0, 0)
    Debug.Print "Bearing is ", , BearingFromOrigin(-5, -2, 0, 0)
     Debug.Print
    Debug.Print "Bearing is ", BearingFromOrigin(-5, 0, 0, 0)
    Debug.Print
    Debug.Print "Bearing is ", , BearingFromOrigin(-5, 2, 0, 0)
    Debug.Print "Bearing is  ", BearingFromOrigin(-5, 5, 0, 0)
    Debug.Print "Bearing is ", , BearingFromOrigin(-2, 5, 0, 0)
    Debug.Print
End Sub


Public Function CheckForNESW(ByVal ipOriginX As Long, ByVal ipOriginY As Long, ByVal ipCurrentX As Long, ByVal ipCurrentY As Long) As Long
    
    If (ipOriginX = ipCurrentX) And (ipOriginY = ipCurrentY) Then
    
        CheckForNESW = -2
        Exit Function
        
    End If
           
    If ipOriginX = ipCurrentX Then
    
        CheckForNESW = IIf(ipCurrentY > ipOriginY, 180, 0)
        Exit Function
        
    End If
    
    If ipOriginY = ipCurrentY Then
    
        CheckForNESW = IIf(ipCurrentX > ipOriginX, 90, 270)
        Exit Function
    End If
    
    CheckForNESW = -1

End Function

