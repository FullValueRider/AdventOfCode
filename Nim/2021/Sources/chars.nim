 const 

#    \r, \c	carriage return
#\n, \l	line feed
#\f	form feed
#\t	tabulator
#\v	vertical tabulator
#\\	backslash
#\"	quotation mark
#\'	apostrophe
#\ '0'..'9'+	character with decimal value d; all decimal digits directly following are used for the character
#\a	alert
#\b	backspace
#\e	escape [ESC]
#\x HH	character with hex value HH; exactly two hex digits are allowed
    twNoString*             : string = ""
    twBar*                  : string = "|"
    twComma*                : string = ","
    twPeriod*               : string = "."
    twSpace*                : string = " "
    twHyphen*               : string = "-"
    twColon*                : string = ":"
    twSemiColon*            : string = ";"
    twHash*                 : string = "#"
    twPlus*                 : string = "+"
    twAsterix*              : string = "*"
    twLParen*               : string = "("
    twRParen*               : string = ")"
    twAmp*                  : string = "@"
    twLBracket*             : string = "["
    twRBracket*             : string = "]"
    twLCurly*               : string = "{"
    twRCurly*               : string = "}"
    twPlainDQuote*          : string = "\""
    twPlainSQuote*          : string = "'"
    twLSmartSQuote*         : string = "‘" #ChrW$(145)   ' Alt+0145
    twRSmartSQuote*         : string = "’" #ChrW$(146)   ' Alt+0146
    twLSMartDQuote*         : string = "“" #ChrW$(147)   ' Alt+0147
    twRSmartDQuote*         : string = "”" #ChrW$(148)   ' Alt+0148
    twTab*                  : string = "\t"
    twCrLf*                 : string = "\r\n"
    twLf*                   : string = "\n"
    twCr*                   : string = "\c"
    twNBsp*                 : string = "\255" #Chr$(255)