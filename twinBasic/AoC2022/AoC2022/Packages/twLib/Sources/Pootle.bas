

Sub TrailTest()

    Dim myT As Trail = Trail.Deb(Coord.Deb(0, 0))
    Dim myT2 As Trail = Trail.Deb(Coord.Deb(0, 0))
    Debug.Print myT.ToString, myT2.ToString
    myT.Move "N", 5
    Debug.Print myT.ToString, myT2.ToString
End Sub