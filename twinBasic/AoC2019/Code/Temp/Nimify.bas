Attribute VB_Name = "Nimify"
Option Explicit
'@IgnoreModule

Public Sub ConvertToNim()

    With ActiveDocument.Styles.Item("Normal")
        
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .ParagraphFormat.SpaceAfter = 0
        .Font.Name = "Consolas"
        .Font.Size = 11
        .NoSpaceBetweenParagraphsOfSameStyle = _
        False
    End With

    
' Manage Start of block end lines
    DoFind "End Sub", vbNullString
    DoFind "End Function", vbNullString
    DoFind "End Property", vbNullString
    DoFind "End If", vbNullString
    DoFind "End With", vbNullString
    DoFind "End Select", vbNullString
    DoFind "End Enum", vbNullString
    DoFind "End Type", vbNullString
    DoFind "Next", vbNullString
    DoFind "Wend", vbNullString
    DoFind "Exit Sub", "return"
    DoFind "Exit Function", "return"
    DoFind "Exit Procedure", "return"
    DoFind "Exit For", "break"
    DoFind "Exit Do", "break"
    'DoFind "Exit While", "break"
    
    DoFind "(Type )(*)(13)", "type^p^tstruct = \2"
    DoFind "(Dim)(*)(^13)", vbNullString
    DoFind "ByVal ", vbNullString
    DoFind "ByRef ", vbNullString
    DoFind "( {1,})As ", ":\1 "
    DoFind "Private ", vbNullString
    DoFind "Public ", vbNullString
    DoFind "(Enum)( )(*)(^013)", "type^p    \3 = enum\4"
    DoFind "(Property Set )(*)([(])", "proc `\2=`(Me:  <object>,"
    DoFind "(Property Let )(*)([(])", "proc `\2=`(Me:  <object>,"
    DoFind "(Property Get )(*)([(])", "proc \2("
    DoFind "Set ", vbNullString
    
    DoFind "Const ", "const "
    DoFind "Do While ", "while "
    DoFind "Do Until ", " while !"
    DoFind "Do^13 ", "while <needs loop condition from below>"
    DoFind "Loop Until ", "//Move condition to the matching while"
    DoFind "Loop While ", "//Move condition to the matching while"
    DoFind "Loop", "while <needs do condition from above>"
    DoFind " Then", ":"
    DoFind "ElseIf", "elif"
    DoFind "Case Else", "else:"
    DoFind "Else", "else:"
    DoFind "If ", "if "
    DoFind "Sub", "proc"
    DoFind "Function", "proc"
    DoFind "For Each", "for"
    DoFind "(For)( )(*)( )(=)( )(*)( To )(*)", "for \3 in \7..\9:"
    DoFind " Step ", "// use countdown format instead."
    DoFind "(Select Case )(*)(^13)", "case \2:\3"
    DoFind "Case ", "of "
    DoFind " To ", ".."

'opeands
    DoFind " [<][>] ", " != "
    DoFind "(IIf[(])(*)(,)(*)(,)(*)([)])", "if \2:\4 else:\6"
    DoFind "(if )(*)( = )(*)(:)", "\1\2 == \4\5"
    DoFind "(elif )(*)( = )(*)(:)", "\1\2 == \4\5"
    DoFind "(while )(*)( = )", "\1\2 == "
    DoFind " Not ", " !"
    DoFind " And ", " and "
    DoFind " Or ", " or "
    DoFind " Xor ", " xor "
    
    DoFind "Debug.Print", "echo fmt"
    
    DoFind " Long", " int32"
    DoFind " Integer", " int16"
    DoFind " Currency", " int64"
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


