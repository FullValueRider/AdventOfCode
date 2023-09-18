Attribute VB_Name = "Layout"
Attribute VB_Description = "Enables field and layout substitution in strings"
Option Explicit
'@Folder("Layout")
'@ModuleDescription("Enables field and layout substitution in strings")

' This module enables fields within strings which represent variables or
' formatting instructions
'
' Variables are indicated by {x} where x is a positive integer.
' e.g.
'       Layout.Format("this string {0} {1}", "Hello", 9)
'
' gives 'this string Hello 9'
'
' Layout fields are of the form {zzx]
' where zz can be
'       nl = new line
'       nt = newline followed by a tab
'       tb = tab
'       x  = the number of times a formatting character is repeated.
'
' If no 'x' is provided then a single layout character is used
' For 'nt' the 'x' refers to the number of newlines.  Only a single tab is inserted
'

'12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Const NEWLINES                                  As String = "{nl"
Private Const TABS                                      As String = "{tb"
Private Const NEWLINES_TAB                              As String = "{nt"

Private Const FIELD_COUNT_NONE                          As String = "}"
Private Const FIELD_COUNT_0                             As String = "0}"
Private Const FIELD_COUNT_1                             As String = "1}"
Private Const LAYOUT_FIELD_TAGS                         As String = NEWLINES & "," & TABS & "," & NEWLINES_TAB

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Primitives (Subs and functions that only depend on VBA or external object libraries).
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Public Function GetLayoutFieldTemplatesArray() As Variant

    GetLayoutFieldTemplatesArray = Split(LAYOUT_FIELD_TAGS, ",")
    
End Function

'@Description("Replace formatting fields of the form '{XX0}' with vbNullString")
Public Function ConvertFormatFieldXX0ToVbNullString(ByVal ipLayoutTemplate As String) As String
Attribute ConvertFormatFieldXX0ToVbNullString.VB_Description = "Replace formatting fields of the form '{XX0}' with vbNullString"

    Dim myLayoutTemplate                As String
    Dim myItem                          As Variant

    myLayoutTemplate = ipLayoutTemplate
    
    For Each myItem In GetLayoutFieldTemplatesArray
    
        myLayoutTemplate = VBA.Replace(myLayoutTemplate, myItem & FIELD_COUNT_0, vbNullString)
        
    Next

    ConvertFormatFieldXX0ToVbNullString = myLayoutTemplate
    
End Function

'@Description("Convert "XX" Layout.Format fields to {XX1}")
'@Ignore AssignedByValParameter
Public Function ConvertFormatFieldXXToXX1(ByVal ipFormatTemplate As String) As String

    Dim myItem                         As Variant

    For Each myItem In GetLayoutFieldTemplatesArray
    
        ipFormatTemplate = VBA.Replace(ipFormatTemplate, myItem & FIELD_COUNT_NONE, myItem & FIELD_COUNT_1)
        
    Next
      
    ConvertFormatFieldXXToXX1 = ipFormatTemplate
      
End Function

'@Description("For format field {XXn}ops the value of n as a Long")
Public Function GetFieldRepeatCount(ByRef ipFormatTemplate As String, ByVal ipFormatField As String) As Long
Attribute GetFieldRepeatCount.VB_Description = "For format field {XXn}ops the value of n as a Long"

    Dim myFormatFieldRepeatLocation                       As Long
    Dim myRepeatCount                          As String

    myFormatFieldRepeatLocation = InStr(ipFormatTemplate, ipFormatField) + Len(ipFormatField)
    
    Do While Mid$(ipFormatTemplate, myFormatFieldRepeatLocation, 1) Like "#"
        
        myRepeatCount = myRepeatCount & Mid$(ipFormatTemplate, myFormatFieldRepeatLocation, 1)
        myFormatFieldRepeatLocation = myFormatFieldRepeatLocation + 1
        
    Loop
    
    GetFieldRepeatCount = CLng(myRepeatCount)
    
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Dependants (Subs and FUnctions that depend on other subs and function in the module/project)
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@Description("Converts string with variable and layout fields to text"}
Public Function Fmt(ByVal ipFormatTemplate As String, Optional ByVal ipArgArray As Variant) As String
    
    Dim myArgArray                                  As Variant
    Dim myReturn                                    As String
    
    If IsMissing(ipArgArray) Then
    
        Fmt = ipFormatTemplate
        Exit Function
        
    End If
    
    myArgArray = HelperForVariants.EnsureArray(ipArgArray)
    '@Ignore AssignmentNotUsed
    myReturn = Layout.ConvertFormatFieldXX0ToVbNullString(ipFormatTemplate)
    myReturn = Layout.ConvertFormatFieldXXToXX1(myReturn)
    myReturn = Layout.ConvertFormatFieldXXnToCharacters(myReturn)
    myReturn = Layout.ConvertVariableFieldsItemingRepresentations(myReturn, myArgArray)
    Fmt = myReturn

End Function

'@Description("Convert {XXn} layout field to fomatting characters
Public Function ConvertFormatFieldXXnToCharacters(ByRef ipFormatTemplate As String) As String

    Dim myItem                                          As Variant
    Dim myReplace                                       As String
    Dim myField                                         As String
    Dim myCount                                         As Long

    For Each myItem In GetLayoutFieldTemplatesArray
            
        If InStr(ipFormatTemplate, myItem) > 0 Then
        
            myCount = GetFieldRepeatCount(ipFormatTemplate, myItem)
            myField = myItem & CStr(myCount) & "}"
            myReplace = GetFormattingReplaceString(myItem, myCount)
            ipFormatTemplate = VBA.Replace(ipFormatTemplate, myField, myReplace)
            
        End If
    
    Next
            
    ConvertFormatFieldXXnToCharacters = ipFormatTemplate
    
End Function

'@Description("Returns a string of formatting characters in line with the formatting tag")
Public Function GetFormattingReplaceString(ByVal ipFormatString As String, ByVal ipRepeatCount As Long) As String
Attribute GetFormattingReplaceString.VB_Description = "Returns a string of formatting characters in line with the formatting tag"
    
    Dim myReturn                                As String

    Select Case ipFormatString
    
    Case NEWLINES
        
        myReturn = VBA.String$(ipRepeatCount, vbCrLf)
            
    Case TABS
        
        myReturn = VBA.String$(ipRepeatCount, vbTab)
            
    Case NEWLINES_TAB
        
        myReturn = VBA.String$(ipRepeatCount, vbCrLf) & vbTab
            
    End Select
    
    GetFormattingReplaceString = myReturn
    
End Function

'@Description("Replace each ocurrence of '{<number>}' with the corresponding stringified item from the parameters list")
Public Function ConvertVariableFieldsItemingRepresentations _
       ( _
       ByVal ipFormatTemplate As String, _
       ByVal ipArgArray As Variant _
       ) As String
Attribute ConvertVariableFieldsItemingRepresentations.VB_Description = "Replace each ocurrence of '{<number>}' with the corresponding stringified item from the parameters list"

    Const GENERIC_VARIABLE_FIELD                            As String = "{XX}"
    Const GENERIC_VARIABLE                                  As String = "XX"

    Dim myField                                             As String
    Dim myIndex                                             As Variant
    Dim myArgArray                                          As Variant
    Dim myReturn                                            As String

    myArgArray = EnsureArray(ipArgArray)
    myReturn = ipFormatTemplate
    
    For myIndex = 0 To UBound(myArgArray)
       
        myField = Replace(GENERIC_VARIABLE_FIELD, GENERIC_VARIABLE, CStr(myIndex))
        myReturn = _
                 VBA.Replace _
                 ( _
                 myReturn, _
                 myField, _
                 HelperForStrings.Stringify(Array(myArgArray(myIndex))) _
                 )
        
    Next
    
    ConvertVariableFieldsItemingRepresentations = myReturn
    
End Function



