Attribute VB_Name = "Pootle"
Option Explicit

Sub TestPermutations()
Debug.Print
    Dim myS As seqC
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(Array("One", "Two", "Three"), Array(10, 20, 30))
    Set myS = Permutations.ByKey(myK)
    fmt.dbg "{0}", myS
End Sub


