VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule ProcedureCanBeWrittenAsFunction
'@PredeclaredId
'@Exposed
'@Folder("Days")
Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C


Private Const TODAY             As String = "\Day10.txt"

Private Type State

    Data                            As String
    ArrS()                          As Long
    ArrNew()                        As Long
    
End Type

Private s                           As State


Public Sub Execute()
    
        Part01
        Part02
        
End Sub

Private Sub Part01()

    Initialise
    
    'Dim myS As String = s.Data
    
    Dim myCount As Long: myCount = 1
    Do Until myCount > 40
        'Debug.Print myCount
        Dim myNewIndex As Long: myNewIndex = 0
        Dim mySIndex As Long: mySIndex = 0
        Dim myCtr As Long: myCtr = 0
        Do Until s.ArrS(mySIndex) = 0
            
            Dim myCurrentchar As Long
            Dim myChar As Long: myChar = s.ArrS(mySIndex)
            
            If myCurrentchar = 0 Then
                myCurrentchar = myChar
                myCtr = 1
            ElseIf myChar = myCurrentchar Then
                myCtr = myCtr + 1
            Else
                s.ArrNew(myNewIndex) = myCtr
                s.ArrNew(myNewIndex + 1) = myCurrentchar
                myCurrentchar = myChar
                myCtr = 1
                myNewIndex = myNewIndex + 2
            End If
            
            mySIndex = mySIndex + 1
        Loop
        
        s.ArrNew(myNewIndex) = myCtr
        s.ArrNew(myNewIndex + 1) = myCurrentchar
        ClearArray s.ArrS
        s.ArrS = s.ArrNew
        'Print10
        ClearArray s.ArrNew
        'Print10
        myCurrentchar = 0
        myNewIndex = myNewIndex + 2
        myCount = myCount + 1
    Loop
    
    Dim myResult As Long: myResult = myNewIndex
    
    fmt.Dbg "The answer to Day {0} part 1 is {1}.  Found is {2}", VBA.Mid$(TODAY, 5, 2), "252594", myResult
    
End Sub

Private Sub Part02()
    
    Initialise
    
    Dim myCount As Long: myCount = 1
    Do Until myCount > 50
       ' Debug.Print myCount
        Dim myNewIndex As Long: myNewIndex = 0
        Dim mySIndex As Long: mySIndex = 0
        Dim myCtr As Long: myCtr = 0
        Do Until s.ArrS(mySIndex) = 0
            
            Dim myCurrentchar As Long
            Dim myChar As Long: myChar = s.ArrS(mySIndex)
            
            If myCurrentchar = 0 Then
                myCurrentchar = myChar
                myCtr = 1
            ElseIf myChar = myCurrentchar Then
                myCtr = myCtr + 1
            Else
                s.ArrNew(myNewIndex) = myCtr
                s.ArrNew(myNewIndex + 1) = myCurrentchar
                myCurrentchar = myChar
                myCtr = 1
                myNewIndex = myNewIndex + 2
            End If
            
            mySIndex = mySIndex + 1
        Loop
        
        s.ArrNew(myNewIndex) = myCtr
        s.ArrNew(myNewIndex + 1) = myCurrentchar
        ClearArray s.ArrS
        s.ArrS = s.ArrNew
        'Print10
        ClearArray s.ArrNew
        'Print10
        myCurrentchar = 0
        myNewIndex = myNewIndex + 2
        myCount = myCount + 1
    Loop
    
    Dim myResult As Long: myResult = myNewIndex
    
    fmt.Dbg "The answer to Day {0} part 2 is {1}.  Found is {2}", VBA.Mid$(TODAY, 5, 2), "3579328", myResult

End Sub



' Process the incoming data to a suitable format
Private Sub Initialise()

    s.Data = Filer.GetFileAsString(AoCRawData & Year & TODAY)
    ReDim s.ArrS(4000000)
    ReDim s.ArrNew(4000000)
    
    Dim myIndex As Long
    For myIndex = 1 To VBA.Len(s.Data)
        s.ArrS(myIndex - 1) = VBA.CLng(VBA.Mid$(s.Data, myIndex, 1))
    Next
    
End Sub

Private Sub ClearArray(ByRef ipArray() As Long)

    Dim myIndex As Long
    For myIndex = 0 To 4000000
        ipArray(myIndex) = 0
    Next
End Sub

Private Sub Print10()
    Dim myS As String
    Dim myIndex As Long
    For myIndex = 0 To 9
        myS = myS & CStr(s.ArrS(myIndex))
    Next
    
    Debug.Print myS
End Sub

