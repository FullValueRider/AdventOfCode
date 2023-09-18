
#Nim escape character sequences
#    \r, \c    carriage return
#\n, \l    line feed
#\f   form feed
#\t   tabulator
#\v   vertical tabulator
#\\   backslash
#\'   quotation mark
#\'   apostrophe
#\ '0'..'9'+    character with decimal value d; all decimal digits directly following are used for the character
#\a    alert
#\b    backspace
#\e    escape [ESC]
#\x HH    character with hex value HH; exactly two hex digits are allowed

#const    twNochar*             : char = ''
const    twHat*                  : char = '^'
const    twBar*                  : char = '|'
const    twComma*                : char = ','
const    twPeriod*               : char = '.'
const    twSpace*                : char = ' '
const    twHyphen*               : char = '-'
const    twColon*                : char = ':'
const    twSemiColon*            : char = ';'
const    twHash*                 : char = '#'
const    twPlus*                 : char = '+'
const    twAsterix*              : char = '*'
const    twLParen*               : char = '('
const    twRParen*               : char = ')'
const    twRAngle*               : char = '>'
const    twLAngle*               : char = '<'
const    twAmp*                  : char = '@'
const    twLBracket*             : char = '['
const    twRBracket*             : char = ']'
const    twLCurly*               : char = '{'
const    twRCurly*               : char = '}'
const    twPlainDQuote*          : char = '\"'
const    twPlainSQuote*          : char = '\''
# const    twLSmartSQuote*         : char = '\‘' #ChrW$(145)   ' Alt+0145
# const    twRSmartSQuote*         : char = '\’' #ChrW$(146)   ' Alt+0146
# const    twLSMartDQuote*         : char = '\“' #ChrW$(147)   ' Alt+0147
# const    twRSmartDQuote*         : char = '\”' #ChrW$(148)   ' Alt+0148
const    twTab*                  : char = '\t'
const    twCrLf*                 : char = '\n'
const    twLf*                   : char = '\n'
const    twCr*                   : char = '\c'
const    twNBsp*                 : char = '\xFF' #Chr$(255)