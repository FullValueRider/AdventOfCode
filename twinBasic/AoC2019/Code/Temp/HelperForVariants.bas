Attribute VB_Name = "HelperForVariants"
'@Folder("VBASupport")
Option Explicit

'12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Function EnsureArray(ByRef ipArgs As Variant) As Variant

    If IsArray(ipArgs) Then
    
        EnsureArray = ipArgs
        
    Else
    
        EnsureArray = Array(ipArgs)
        
    End If
    
End Function

'@Description("Replacement for the 'IsEmpty' function which works with variants containing arrays")
Public Function VariantIsEmpty(ByRef ipVariant As Variant) As Boolean
Attribute VariantIsEmpty.VB_Description = "Replacement for the 'IsEmpty' function which works with variants containing arrays"
    ' A variant containing an unPredeclaredIdConstructionStatus array can never be empty unlike a
    ' containing a primitive type.  Thus we need an alternative way for defining
    ' empty as an unPredeclaredIdConstructionStatus array.  Ubound on an uninitialed array generates
    ' a VBA error rather thanoping 0 as you might expect.

    Const ARRAY_TYPE                                As String = "()"

    '@Ignore VariableNotUsed
    Dim myDummy                                    As Long

    If VBA.InStr(VBA.TypeName(ipVariant), ARRAY_TYPE) = 0 Then
    
        VariantIsEmpty = VBA.IsEmpty(ipVariant)
        
    Else
    
        On Error Resume Next
        myDummy = UBound(ipVariant)
        VariantIsEmpty = Not (VBA.Err.Number = 0)
        On Error GoTo 0
        
    End If
    
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Dependants (Subs and FUnctions that depend on other subs and function in the module/project)
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

'@Description("Function to complement VariantIsEmpty")
Public Function VariantHasValue(ByRef ipVariant As Variant) As Boolean
Attribute VariantHasValue.VB_Description = "Function to complement VariantIsEmpty"
    
    VariantHasValue = Not VariantIsEmpty(ipVariant)
    
End Function

'@Description("Function to complement the variant 'IsMissing' Method")
Public Function IsNotMissing(ByVal ipVariant As Variant) As Boolean
Attribute IsNotMissing.VB_Description = "Function to complement the variant 'IsMissing' Method"

    '@Ignore IsMissingOnInappropriateArgument
    IsNotMissing = Not IsMissing(ipVariant)
    
End Function
