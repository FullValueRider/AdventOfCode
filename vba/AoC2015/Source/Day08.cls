VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule EmptyCaseBlock
'@PredeclaredId
'@Exposed
'@Folder("Days")
Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========


Private Const TODAY                 As String = "\Day08.txt"

Private Type State

    Data                            As seqC
    
End Type

Private s                           As State


Public Sub Execute()
        Part01
        Part02

End Sub

Private Sub Part01()
'    /* For part1 there are three scenarios to consider
'        \"
'        \\
'        \x??
'    when we find any one of these sequences we replace with a single character
'    Because we are not required to extract information from the in memory strings we can use
'    any single character (e.g. #) to indicate where a replacmeent has taken place
'
'    All input strings are surrounded by double quotes so these need to be ignored when
'    constructing the in memory string
'    */
    
    Initialise
    
    Dim myCodeStrLen As Long
    myCodeStrLen = 0
    
    Dim myInMemStrLen As Long
    myInMemStrLen = 0
    
    Dim myItem As Variant
    Dim myItems As Iteritems: Set myItems = Iteritems(s.Data)
    Do
        myItem = myItems.curItem(0)
        Dim myCodeStr As String
        myCodeStr = myItem
        myCodeStrLen = myCodeStrLen + VBA.Len(myCodeStr)
        
        Dim myInMemStr As String
        myInMemStr = RemoveControlStrings(myCodeStr)
        myInMemStrLen = myInMemStrLen + VBA.Len(myInMemStr)
        
    Loop While myItems.MoveNext
   
    Dim myResult As Long
    myResult = myCodeStrLen - myInMemStrLen
    Debug.Print fmt.Text("The answer to Day {0} part 1 is 1333.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult)
    
End Sub

Private Sub Part02()

    ' for part 2 the escape encodinga are
    ' " becomes

    Initialise
    
    Dim myCodeStrLen As Long
    myCodeStrLen = 0
    
    Dim myInMemStrLen As Long
    myInMemStrLen = 0
    
    Dim myItem As Variant
    Dim myItems As Iteritems: Set myItems = Iteritems(s.Data)
    Do
        myItem = myItems.curItem(0)
        Dim myCodeStr As String
        myCodeStr = myItem
        myCodeStrLen = myCodeStrLen + VBA.Len(myCodeStr)
        
        Dim myInMemStr As String
        myInMemStr = EnhanceControlStrings(myCodeStr)
        myInMemStrLen = myInMemStrLen + VBA.Len(myInMemStr)
        
    Loop While myItems.MoveNext
   
    Dim myResult As Long
    myResult = myInMemStrLen - myCodeStrLen
    fmt.Dbg "The answer to Day {0} Part 02 is 2046.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult

End Sub


Private Function RemoveControlStrings(ByRef ipString As String) As String

    Dim myInMemStr As String
    myInMemStr = vbNullString
    
    Dim myIndex As Long
    myIndex = 1
    Do While myIndex <= VBA.Len(ipString)
    
        Dim myChar As String
        myChar = VBA.Mid$(ipString, myIndex, 1)
      
        Select Case myChar
        
            Case chars.twPlainDQuote
            
                'ignore the character, i.e. do nothing
            
            Case chars.twBSlash
            
                ' we have found the start of an escape sequence
                myIndex = myIndex + 1
                myChar = VBA.Mid$(ipString, myIndex, 1)
                
                Select Case myChar
                
                    Case chars.twPlainDQuote
                        
                        'replace \" with a single character '#'
                        myInMemStr = myInMemStr & "#"
                    
                    Case "x"
                        'found a \x??
                        
                        myIndex = myIndex + 2
                        myInMemStr = myInMemStr & "#"
                    
                    
                    Case chars.twBSlash
                        'replace \\ with a single character '#'
                        myInMemStr = myInMemStr & "#"
                        
                    Case Else
                        ' not a control sequence so \x is two chars
                        myInMemStr = myInMemStr & "##"
                    
                End Select
                
            Case Else
            
                myInMemStr = myInMemStr & myChar
            
        End Select
        
        myIndex = myIndex + 1
        
    Loop
    
    RemoveControlStrings = myInMemStr
    
End Function

Private Function EnhanceControlStrings(ByRef ipString As String) As String

    Dim myInMemStr As String
    myInMemStr = vbNullString
    
    Dim myIndex As Long
    myIndex = 1
    Do While myIndex <= VBA.Len(ipString)
    
        Dim myChar As String
        myChar = VBA.Mid$(ipString, myIndex, 1)
      
        Select Case myChar

            Case chars.twBSlash
                
                Dim myNextchar As String
                myNextchar = VBA.Mid$(ipString, myIndex + 1, 1)
                
                Select Case myNextchar

                    Case chars.twPlainDQuote

                        myInMemStr = myInMemStr & "####"
                        myIndex = myIndex + 1

                    Case chars.twBSlash

                        myInMemStr = myInMemStr & "####"
                        myIndex = myIndex + 1
                    Case "x"

                        myInMemStr = myInMemStr & "##"

                    Case Else

                        myInMemStr = myInMemStr & "#"

                End Select

            Case chars.twPlainDQuote

                myInMemStr = myInMemStr & "###"


            Case Else

                myInMemStr = myInMemStr & myChar

        End Select
        
        myIndex = myIndex + 1
        
    Loop
    
    EnhanceControlStrings = myInMemStr
    
End Function


Private Sub Initialise()

    Set s.Data = seqC.Deb(Filer.GetFileAsArrayOfStrings(AoCRawData & Year & TODAY))
   
End Sub
