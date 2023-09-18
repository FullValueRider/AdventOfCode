Attribute VB_Name = "HelperForStrings"
'@Folder("StringyStuff")
Option Explicit


'12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

'@Description("True if ipValue is a OStr")
Public Function IsOStr(ByRef ipValue As Variant) As Boolean
Attribute IsOStr.VB_Description = "True if ipValue is a OStr"

    IsOStr = InStr(TypeName(ipValue), "StrO") > 0
    
End Function

'@Description("True if ipVariant is a String or OStr")
Public Function IsStringy(ByRef ipVariant As Variant) As Boolean
Attribute IsStringy.VB_Description = "True if ipVariant is a String or OStr"

    IsStringy = InStr("String,StrO", TypeName(ipVariant)) > 0
    
End Function

Public Function IsAString(ByRef ipVariant As Variant) As Boolean
    IsAString = VarType(ipVariant) = vbString
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Dependants (Subs and Functions that depend on other subs and function in the module/project)
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@Description("Converts the parameter ip_arg to a single string.  ipArgArray must be a variant containing an array")
Public Function Stringify _
       ( _
       ByVal ipArgArray As Variant, _
       Optional ByVal ipSeparator As String = vbNullString, _
       Optional ByVal ipBracketArrays As Boolean = True _
       ) As String
Attribute Stringify.VB_Description = "Converts the parameter ip_arg to a single string.  ipArgArray must be a variant containing an array"

    ' Attempts to convert each argument to a string representation
    ' substitues an error string if no conversion is possible
    ' e.g. when we encounter an object
    ' If an argument is itDebutante an array then the strigified argument
    ' will be placed between 'vbcrlf[' and ']' and each item in the array
    ' will be seperated by commas

    Dim myArg                                       As Variant
    Dim myReturn                                    As String
    Dim myArgArray                                  As Variant

    myArgArray = ipArgArray
    
    For Each myArg In myArgArray
            
        myReturn = IIf(Len(myReturn), myReturn & ipSeparator, myReturn)
        
        If IsArray(myArg) Then
            
            myReturn = myReturn & StringifyArray(myArg, ipBracketArrays)
            
        Else
        
            myReturn = myReturn & StringifyItem(myArg)
            
        End If
        
    Next
    
    Stringify = myReturn
      
End Function

Public Function StringifyItem(ByRef ipItem As Variant) As String

    Dim myReturn                       As String
    
    ' Stringifiable objects should be identified here and stringified
    Select Case TypeName(ipItem)
    
    Case "StrO"
        
        
        myReturn = ipItem.GetValue
    
    Case Else
    
        On Error Resume Next
        myReturn = CStr(ipItem)
    
        'IsBad
        If VBA.Err.Number > 0 Then
    
            On Error GoTo 0
            myReturn = "{Can't stringify Type: " & TypeName(ipItem) & "}"
        
        End If
            
        On Error GoTo 0
            
    End Select
    
    StringifyItem = myReturn
      
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Public Function StringifyArray _
       ( _
       ByVal ipArgArray As Variant, _
       Optional ByVal ipPutSquareBracketsAroundArrayFlag As Boolean = True, _
       Optional ByVal ipSeparatorBetweenArrayItems As String = "," _
       ) As String

    Dim myArg                                       As Variant
    Dim myArgArray                                  As Variant
    Dim myReturn                                    As String
    Dim myLeftSquareBracket                         As String
    Dim myRightSquareBracket                        As String
    Dim mySeparatorBetweenArrayItems                As String
    
    myArgArray = ipArgArray
    myLeftSquareBracket = IIf(ipPutSquareBracketsAroundArrayFlag, "[", vbNullString)
    myRightSquareBracket = IIf(ipPutSquareBracketsAroundArrayFlag, "]", vbNullString)
    myReturn = myLeftSquareBracket
    
    For Each myArg In myArgArray
    
        mySeparatorBetweenArrayItems = IIf(myReturn = myLeftSquareBracket, vbNullString, ipSeparatorBetweenArrayItems)
    
        If IsArray(myArg) Then
        
            
            myReturn = myReturn & mySeparatorBetweenArrayItems & StringifyArray(myArg)
            
        Else
        
            myReturn = myReturn & mySeparatorBetweenArrayItems & StringifyItem(myArg)
            
        End If
        
    Next
    
    StringifyArray = myReturn & myRightSquareBracket
    
End Function

'@Description("False if ipVariant is a String or OStr")
Public Function IsNotStringy(ByRef ipVariant As Variant) As Boolean
Attribute IsNotStringy.VB_Description = "False if ipVariant is a String or OStr"

    IsNotStringy = Not IsStringy(ipVariant)
    
End Function

'@Description("False if ipValue is a OStr")
Public Function IsNoItem(ByRef ipValue As Variant) As Boolean
Attribute IsNoItem.VB_Description = "False if ipValue is a OStr"

    IsNoItem = Not IsOStr(ipValue)
    
End Function

Public Function Reverse(ByVal ipString As String) As String
    
    Dim myIndex As Long
    Dim myReturn As String
    
    For myIndex = 1 To Len(ipString)
    
        myReturn = Mid$(ipString, myIndex, 1) & myReturn
        
    Next
        
    Reverse = myReturn
    
End Function


