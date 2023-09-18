VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dAY12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Days")
 Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Private Const TODAY             As String = "\Day12.txt"

Private Type State
    Data                            As String
End Type

Private s                           As State


Public Sub Execute()
    Part01
    Part02
End Sub

Private Sub Part01()

    Initialise
    
    Dim myResult As Long: myResult = SumNumbersInString(s.Data)
    
    fmt.Dbg "The answer to Day {0} part 1 is 191164.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult
    
End Sub

Private Sub Part02()

    Initialise
        
    Dim myJson As Object: Set myJson = JsonConverter.ParseJson(s.Data)
    
'    If TypeName(myJson) = "Collection" Then
'        Set myJson = seqc.Deb.Bind(myJson)
'    Else
'        Set myJson = kvpha.Deb.AddPairs(myJson.Keys, myJson.Items)
'    End If
    
    PurgeObjectsWithRedItems myJson
    
    Dim myString As String: myString = JsonConverter.ConvertToJson(myJson)
    
    Dim myResult As Long: myResult = SumNumbersInString(myString)
    
    fmt.Dbg "The answer to Day {0} part 2 is 87842.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult

End Sub

' Process the incoming data to a suitable format
Private Sub Initialise()

    s.Data = Filer.GetFileAsString(AoCRawData & Year & TODAY)
        
End Sub


Public Sub PurgeObjectsWithRedItems(ByRef ipDecodedJson As Variant)

    If VBA.TypeName(ipDecodedJson) = "Collection" Then
        Set ipDecodedJson = seqC.Deb(ipDecodedJson)
    Else
        Set ipDecodedJson = kvpha.Deb.AddPairs(ipDecodedJson.Keys, ipDecodedJson.Items)
    End If
    
    Dim myItems As Iteritems
        
    Select Case True
    
        Case GroupInfo.IsDictionary(ipDecodedJson)
        
            If ipDecodedJson.HoldsItem("red") Then
                ipDecodedJson.RemoveAll ' set ipvar = nothing doesnot work.
                Exit Sub
            End If
            
            
            Set myItems = Iteritems.Deb(ipDecodedJson).MoveToEndIndex
            Do
                PurgeObjectsWithRedItems myItems.curItem(0)
            Loop While myItems.MovePrev
            
            
        Case GroupInfo.IsIndexFrom1(ipDecodedJson)
        
            Set myItems = Iteritems.Deb(ipDecodedJson).MoveToEndIndex
            Do
                PurgeObjectsWithRedItems myItems.curItem(0)
            Loop While myItems.MovePrev
            
            
    End Select

'     Red will be an Item in a dictionary
'     so we can ignore everything except dictionaries
'    Debug.Print TypeName(ipDecodedJson)
'    Dim myItems As IterItems
'    Set myItems = IterItems(ipDecodedJson)
'    Do
'        Debug.Print TypeName(myItems.Item(0))
'        If VBA.IsObject(myItems.Item(0)) Then
'            Debug.Print "-" & TypeName(myItems.Item(0))
'            If GroupInfo.IsDictionary(myItems.Item(0)) Then
'                Dim myK As Kvpc: Set myK = Kvpc.Deb.AddPairs(seqc(myItems.Item(0).Keys), seqc(myItems.Item(0).Items))
'                If myK.HoldsItem("red") Then
'                    Dim myD As Dictionary: Set myD = myItems.Item(0)
'
'                    myD.RemoveAll ' set ipvar = nothing doesnot work.
'                    Exit Sub
'                Else
'                    PurgeObjectsWithRedItems myItems.Item(0)
'                End If
'
'            End If
'        End If
'
'    Loop While myItems.MoveNext
        
End Sub

Private Function SumNumbersInString(ByVal ipString As String) As Long

    Dim myResult As Long: myResult = 0
    Dim myNumber As String: myNumber = vbNullString
    Dim myNumbers As seqC: Set myNumbers = seqC(chars.Number)
    
    Dim myChars As Iteritems: Set myChars = Iteritems.Deb(ipString)
    Do
        Dim myChar As String
        myChar = myChars.curItem(0)
        
        If myChar = "-" Then
        
            myNumber = myNumber & myChar
            
        ElseIf myNumbers.LacksItem(myChar) Then 'Or myChar = "-"
        
            If VBA.Len(myNumber) > 0 Then
       
                myResult = myResult + VBA.CLng(myNumber)
                myNumber = vbNullString
                
            End If
            
        Else
        
            myNumber = myNumber & myChar
            
        End If
        
    Loop While myChars.MoveNext
    
    SumNumbersInString = myResult
    
End Function

