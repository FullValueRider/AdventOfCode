Attribute VB_Name = "TestingDay05"
Option Explicit
'@Ignoremodule
Private Sub Day05Testing()

    

    Dim myNice As String
    Dim myWord As String
    Debug.Print "Nice V1"
    myWord = "ugknbfddgicrmopn"
    myNice = "naughty"
    If Day05.IsNiceV1(myWord) Then myNice = "nice"
    Debug.Print Layout.Fmt("{0} is nice: {1}", myWord, myNice)

    myWord = "aaa"
    myNice = "naughty"
    If Day05.IsNiceV1(myWord) Then myNice = "nice"
    Debug.Print Layout.Fmt("{0} is nice: {1}", myWord, myNice)

     myWord = "jchzalrnumimnmhp"
    myNice = "naughty"
    If Day05.IsNiceV1(myWord) Then myNice = "nice"
    Debug.Print Layout.Fmt("{0} is naughty: {1}", myWord, myNice)

    myWord = "haegwjzuvuyypxyu"
    myNice = "naughty"
    If Day05.IsNiceV1(myWord) Then myNice = "nice"
    Debug.Print Layout.Fmt("{0} is naughty: {1}", myWord, myNice)

    myWord = "dvszwmarrgswjxmb"
    myNice = "naughty"
    If Day05.IsNiceV1(myWord) Then myNice = "nice"
    Debug.Print Layout.Fmt("{0} is naughty: {1}", myWord, myNice)
    Debug.Print
    Debug.Print "Nice V2"
    myWord = "qjhvhtzxzqqjkmpb"
    myNice = "naughty"
    If Day05.IsNiceV2(myWord) Then myNice = "nice"
    Debug.Print Layout.Fmt("{0} is nice: {1}", myWord, myNice)


    myWord = "xxyxx"
    myNice = "naughty"
    If Day05.IsNiceV2(myWord) Then myNice = "nice"
    Debug.Print Layout.Fmt("{0} is nice: {1}", myWord, myNice)

    myWord = "uurcxstgmygtbstg"
    myNice = "naughty"
    If Day05.IsNiceV2(myWord) Then myNice = "nice"
    Debug.Print Layout.Fmt("{0} is naughty: {1}", myWord, myNice)

    myWord = "ieodomkazucvgmuy"
    myNice = "naughty"
    If Day05.IsNiceV2(myWord) Then myNice = "nice"
    Debug.Print Layout.Fmt("{0} is naughty: {1}", myWord, myNice)

End Sub

