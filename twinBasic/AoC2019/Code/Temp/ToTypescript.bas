Attribute VB_Name = "ToTypescript"
Option Explicit

'@IgnoreModule
Public Sub ConvertToTypeScript()

    With ActiveDocument.Styles.Item("Normal")
        
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .ParagraphFormat.SpaceAfter = 0
        .Font.Name = "Consolas"
        .Font.Size = 8
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
    DoFind "Wend", "}"
    DoFind "Exit Sub", "}"
    DoFind "Exit Function", "}"
    DoFind "Exit Procedure", "}"
    DoFind "Exit For", "break"
    DoFind "Exit Do", "break"
    'DoFind "Exit While", "break"
    DoEvents
    
    DoFind "(Type)( )(*)(^13)", "interface \3^p{^p"
    DoFind "(Dim)(*)( As )(*)(^13)", "var \2: \4^p"
    DoFind "ByVal ", vbNullString
    DoFind "ByRef ", vbNullString
    DoFind "( As )", ": "
    DoFind "Private ", vbNullString
    DoFind "Public ", "export "
    DoFind "(Enum)( )(*)(^13)", " enum \3^p{^p"
    DoFind "(Property Set )(*)([(])(*)(^13)(*)(Set )(*)( = )", "set \2\3\4^p{^p\6this.\8 =  "
    DoFind "(Property Let )(*)([(])(*)(^13)(*)( {2,})([sp.]{2,2})(\2)", "set \2\3\4^p{^p\6\7this.\8\9 "
    DoFind "(Property Let )(*)([(])(*)(^13)", "set \2\3\4^p"
    DoFind "(Property Get )(*)([(])(*)(^13)(*)(Set )(\2)( = )", "get \2\3\4^p{^p\6return "
    DoFind "(Property Get )(*)([(])(*)(^13)(*)( {2,})(\2)( = )", "get \2\3\4^p{^p\6\7return "
    DoFind "Set ", vbNullString
    
    DoEvents
    DoFind "Const ", "const "
    DoFind "(Do While)(*)(^13)", "while (\2)^p{"
    DoFind "(Do Until)(*)(^13)", " while (!\2)^p{"
    DoFind "Do^13 ", "while <needs loop condition from below>"
    DoFind "(Loop Until)(*)(^13)", "} while(!\2)^p"
    DoFind "(Loop While)(*)(^13)", "} while(\2)^p"
    DoFind "Loop", "}"
    DoFind "(ElseIf)(*)(Then)", "}^pelse if (\2)^p{"
    DoFind "Case Else", "^tbreak;^pdefault:"
    DoFind "Else", "}^pelse^p{"
    DoFind "(Sub)(*)(^13)", "function \2^p{"
    DoFind "(Function)(*)(^13)", "function \2^p{"
    DoFind "(For Each)(*)(in)(*)(^13)", "for (var \1 of \4)^p{"
    DoFind "(For )(*)(=)(*)(To)(*)( Step )-(*)(^13)", "for (\2=\4; \2<=\6 ; \2-=\9)^p{"
    DoFind "(For )(*)(=)(*)(To)(*)( Step )(*)(^13)", "for (\2=\4; \2<=\6 ; \2+=\8)^p{"
    DoFind "(For )(*)( = )(*)(To)(*)(^13)", "for (\2=\4; \2<=\6 ; \2++)^p{"
    'DoFind " Step ", "// use countdown format instead."
    DoFind "(Select Case )(*)(^13)", "switch (\2)^p{"
    DoFind "(^13)( {1,})(Case )(*)(: )", "\2^p^tbreak;^p\2}^p\2case \4:^p\2{^p"
    DoFind " To ", ".."
    'Stop
    'DoFind "(switch)(*)(^13)(*)([}])", "\1\2^p{"
    DoEvents
'opeands
    DoFind " [<][>] ", " != "
    DoFind "(IIf[(])(*)(,)(*)(,)(*)([)])", "if \2 {\4} else {\6}"
    DoFind "(If )(*)(Then)", "if (\2)^p{"
    DoFind "(if )(*)( = )(*)(^13)", "\1\2 == \4\5"
    DoFind "(else if )(*)( = )(*)(^13)", "\1\2 == \4\5"
    DoFind "(while )(*)( = )", "\1\2 == "
    DoFind " Not ", " !"
    DoFind " And ", " && "
    DoFind " Or ", " || "
    DoFind " Xor ", " xor "
    
    DoFind "(Debug.Print)(*)(^13)", "console.log(\2);^p"
    DoEvents
   
    DoFind " Integer", " number"
    DoFind " LongLong", " number"
    DoFind " Long", " number"
    DoFind " String", " string"
    DoFind " Boolean", " bool"
    DoFind " Single", " number"
    DoFind " Double", "number"
    
    ' Manage spacing
    'DoFind "    ^013", vbNullString
    DoFind "(^13)(    )(^13)", "\2\3"
    DoFind "(^13)(^13)(    )", "\2\3"
    DoFind "( {1,})(^13)", "\2"
    DoFind "(')( {1,})", "// "
    DoFind "(')([! ])", "// \2"
    
    
    DoFind "(^13{4,})", "^p^p^p"
End Sub


'Public Sub DoFind(ByVal ipFind As String, ByVal ipReplace As String)
'
'    With ActiveDocument.StoryRanges.Item(wdMainTextStory)
'
'        With .Find
'
'            .ClearFormatting
'            .Replacement.ClearFormatting
'            .Text = ipFind
'            .Replacement.ClearFormatting
'            .Replacement.Text = ipReplace
'            .Wrap = wdFindContinue
'            .MatchWildcards = True
'            .MatchCase = True
'            .Execute Replace:=wdReplaceAll
'
'        End With
'
'    End With
'
'End Sub

