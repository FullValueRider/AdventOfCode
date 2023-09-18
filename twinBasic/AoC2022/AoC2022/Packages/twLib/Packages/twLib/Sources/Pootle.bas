
Public Sub tESTpARAMaRRAY()

    Dim myArray As Variant = Array(10, 20, 30, 40, 50, 60)
    TakeParamArray myArray
End Sub

Sub TakeParamArray(ParamArray ipParamarray() As Variant)

    Dim myTmp As Variant = ipParamarray(0)
    myTmp(3) = 600
    ipParamarray(0)(3) = 100
    

End Sub