Attribute VB_Name = "Rustify"
Option Explicit
'@IgnoreModule
Public Sub ConvertToRust()

    With ActiveDocument.Styles.Item("Normal")
        
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .ParagraphFormat.SpaceAfter = 0
        .Font.Name = "Consolas"
        .Font.Size = 11
        .NoSpaceBetweenParagraphsOfSameStyle = _
        False
    End With

    
' Manage Start of block end lines
    DoFind "End Sub", "}"
    DoFind "End Function", "}"
    DoFind "End Property", "}"
    DoFind "End If", "}"
    DoFind "End With", "}"
    DoFind "End Select", "}"
    DoFind "End Enum", "}"
    DoFind "End Type", "}"
    DoFind "Next", "}"
    DoFind "Wend", vbNullString
    DoFind "Exit Sub", "return"
    DoFind "Exit Function", "return"
    DoFind "Exit Procedure", "return"
    DoFind "Exit For", "break"
    DoFind "Exit Do", "break"
    'DoFind "Exit While", "break"
    
    
    DoFind "(Type )(*)(^13)", "struct \2^p{"
    DoFind "(Dim)(*)(^13)", vbNullString
    DoFind "ByVal ", vbNullString
    DoFind "ByRef ", vbNullString
    DoFind "([)])( As )(*)(^13)", ") -> \3^p{"
    DoFind "( {1,})As ", ":\1 "
    DoFind "Private ", vbNullString
    DoFind "Public ", "pub "
    DoFind "(Enum)( )(*)(^013)", " enum \3^p{^p"
    DoFind "(Property Set )(*)([(])", "fn Set\2=`(&mut self"
    DoFind "(Property Let )(*)([(])", "fn Set\2=`(&mut self"
    DoFind "(Property Get )(*)([(])", "fn Get\2("
    DoFind "Set ", vbNullString
    
    DoFind "Const ", "const "
    DoFind "Do While ", "while "
    DoFind "Do Until ", " while !"
    DoFind "Do^13 ", "while <needs loop condition from below>"
    DoFind "Loop Until ", "//Move condition to the matching while"
    DoFind "Loop While ", "//Move condition to the matching while"
    DoFind "Loop", "while <needs do condition from above>"
    DoFind " Then", "^p{"
    DoFind "ElseIf", "}^pelse if"
    DoFind "Case Else", "_"
    DoFind "Else", "else:"
    DoFind "If ", "if "
    DoFind "Sub", "fn"
    DoFind "Function", "fn"
    DoFind "For Each", "for"
    DoFind "(For)( )(*)( )(=)( )(*)( To )(*)", "for \3 in \7..\9:"
    DoFind " Step ", "// use countdown format instead."
    DoFind "(Select Case )(*)(^13)", "match \2^p{"
    DoFind "(Case)(*)(: )", "\2 => "
    DoFind " To ", ".."

'opeands
    DoFind " [<][>] ", " != "
    DoFind "(IIf[(])(*)(,)(*)(,)(*)([)])", "if \2 {\4} else {\6}"
    DoFind "(if )(*)( = )(*)(^13)", "\1\2 == \4\5"
    DoFind "(else if )(*)( = )(*)(^13)", "\1\2 == \4\5"
    DoFind "(while )(*)( = )", "\1\2 == "
    DoFind " Not ", " !"
    DoFind " And ", " and "
    DoFind " Or ", " or "
    DoFind " Xor ", " xor "
    
    DoFind "(Debug.Print)(*)(^13)", "PrintLn!(\2);^p"
    
    DoFind " Long", " i32"
    DoFind " Integer", " i16"
    DoFind " LongLong", " i64"
    DoFind " String", " string"
    DoFind " Boolean", " bool"
    ' Manage spacing
    'DoFind "    ^013", vbNullString
    DoFind "(^13)(    )(^13)", "\2\3"
    DoFind "(^13)(^13)(    )", "\2\3"
    DoFind "( {1,})(^13)", "\2"
    DoFind "(')( {1,})", "// "
    DoFind "(')([! ])", "// \2"
    
    
    DoFind "(^13{4,})", "^p^p^p"
End Sub


Public Sub DoFind(ByVal ipFind As String, ByVal ipReplace As String)
    
    With ActiveDocument.StoryRanges.Item(wdMainTextStory)
        
        With .Find
        
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ipFind
            .Replacement.ClearFormatting
            .Replacement.Text = ipReplace
            .Wrap = wdFindContinue
            .MatchWildcards = True
            .MatchCase = True
        
            .Execute Replace:=wdReplaceAll
        End With
    
    End With
    
End Sub


