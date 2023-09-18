Attribute VB_Name = "Char"
'@IgnoreModule ConstantNotUsed
Option Explicit

'@Folder("VBALib")
'Readonly values - Pseudo constants
'@Ignore EmptyStringLiteral
'Public Function twNoString() As String: twNoString As String = "": End Function
'Public Function twBar() As String: twBar As String = "|": End Function
'Public Function twComma() As String: twComma As String = ",": End Function
'Public Function twPeriod() As String: twPeriod As String = ".": End Function
'Public Function twSpace() As String: twSpace As String = " ": End Function
'Public Function twHyphen() As String: twHyphen As String = "-": End Function
'Public Function twColon() As String: twColon As String = ":": End Function
'Public Function twSemiColon() As String: twSemiColon As String = ";": End Function
'Public Function twHash() As String: twHash As String = "#": End Function
'Public Function twPlus() As String: twPlus As String = "+": End Function
'Public Function twAsterix() As String: twAsterix As String = "*": End Function
'Public Function twLParen() As String: twLParen As String = "(": End Function
'Public Function twRParen() As String: twRParen As String = ")": End Function
'Public Function twAmp() As String: twAmp As String = "@": End Function
'Public Function twLBracket() As String: twLBracket As String = "[": End Function
'Public Function twRBracket() As String: twRBracket As String = "]": End Function
'Public Function twLCurly() As String: twLCurly As String = "{": End Function
'Public Function twRCurly() As String: twRCurly As String = "}": End Function
'Public Function twPlainDQuote() As String: twPlainDQuote As String = """": End Function
'Public Function twPlainSQuote() As String: twPlainSQuote As String = "'": End Function
'Public Function twLSmartSQuote() As String: twLSmartSQuote As String = ChrW$(145): End Function ' Alt+0145
'Public Function twRSmartSQuote() As String: twRSmartSQuote As String = ChrW$(146): End Function ' Alt+0146
'Public Function twLSMartDQuote() As String: twLSMartDQuote As String = ChrW$(147): End Function ' Alt+0147
'Public Function twRSmartDQuote() As String: twRSmartDQuote As String = ChrW$(148): End Function ' Alt+0148
'Public Function twTab() As String: twTab As String = vbTab: End Function
'Public Function twCrLf() As String: twCrLf As String = vbCrLf: End Function
'Public Function twLf() As String: twLf As String = vbLf: End Function
'Public Function twCr() As String: twCr As String = vbCr: End Function
'Public Function twNBsp() As String: twNBsp As String = Chr$(255): End Function
    
Public Const twEqual                As string = "="
Public Const twLArrow               As string = "<"
Public Const twRArrow               As string = ">"
Public Const twNoString             As String = ""
Public Const twBar                  As String = "|"
Public Const twComma                As String = ","
Public Const twPeriod               As String = "."
Public Const twSpace                As String = " "
Public Const twHyphen               As String = "-"
Public Const twColon                As String = ":"
Public Const twSemiColon            As String = ";"
Public Const twHash                 As String = "#"
Public Const twPlus                 As String = "+"
Public Const twAsterix              As String = "*"
Public Const twLParen               As String = "("
Public Const twRParen               As String = ")"
Public Const twAmp                  As String = "@"
Public Const twLBracket             As String = "["
Public Const twRBracket             As String = "]"
Public Const twLCurly               As String = "{"
Public Const twRCurly               As String = "}"
Public Const twBSlash               As string = "\"
Public Const twFSlash               As string = "/"
Public Const twPlainDQuote          As String = """"
Public Const twPlainSQuote          As String = "'"
Public Const twLSmartSQuote         As String = "‘" 'ChrW$(145)   ' Alt+0145
Public Const twRSmartSQuote         As String = "’" 'ChrW$(146)   ' Alt+0146
Public Const twLSMartDQuote         As String = "“" 'ChrW$(147)   ' Alt+0147
Public Const twRSmartDQuote         As String = "”" 'ChrW$(148)   ' Alt+0148
Public Const twTab                  As String = vbTab
Public Const twCrLf                 As String = vbCrLf
Public Const twLf                   As String = vbLf
Public Const twCr                   As String = vbCr
Public Const twNBsp                 As String = "ÿ" 'Chr$(255)

