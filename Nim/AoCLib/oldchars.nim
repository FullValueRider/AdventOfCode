

#    \r, \c    carriage return
#\n, \l    line feed
#\f   form feed
#\t   tabulator
#\v   vertical tabulator
#\\   backslash
#\"   quotation mark
#\'   apostrophe
#\ '0'..'9'+    character with decimal value d; all decimal digits directly following are used for the character
#\a    alert
#\b    backspace
#\e    escape [ESC]
#\x HH    character with hex value HH; exactly two hex digits are allowed



#const    twNoString*             : string = ""
const    twHat*                  : string = "^"
const    twBar*                  : string = "|"
const    twComma*                : string = ","
const    twPeriod*               : string = "."
const    twSpace*                : string = " "
const    twHyphen*               : string = "-"
const    twColon*                : string = ":"
const    twSemiColon*            : string = ";"
const    twHash*                 : string = "\x23"
const    twPlus*                 : string = "+"
const    twAsterix*              : string = "*"
const    twLParen*               : string = "("
const    twRParen*               : string = ")"
const    twRAngle*               : string = ">"
const    twLAngle*               : string = "<"
const    twAmp*                  : string = "@"
const    twLBracket*             : string = "["
const    twRBracket*             : string = "]"
const    twLCurly*               : string = "{"
const    twRCurly*               : string = "}"
const    twPlainDQuote*          : string = "\""
const    twPlainSQuote*          : string = "'"
const    twLSmartSQuote*         : string = "‘" #ChrW$(145)   ' Alt+0145
const    twRSmartSQuote*         : string = "’" #ChrW$(146)   ' Alt+0146
const    twLSMartDQuote*         : string = "“" #ChrW$(147)   ' Alt+0147
const    twRSmartDQuote*         : string = "”" #ChrW$(148)   ' Alt+0148
const    twTab*                  : string = "\t"
const    twCrLf*                 : string = "\r\n"
const    twLf*                   : string = "\n"
const    twCr*                   : string = "\c"
const    twNBsp*                 : string = "\255" #Chr$(255)