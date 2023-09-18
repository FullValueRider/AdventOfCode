VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
Option Explicit

Private Const TODAY                             As String = "\Day06.txt"

Private Type State

    Display                                     As BulbDisplay
    Data                                        As seqC
    Instructions                                As seqC
    Lights                                      As Variant

End Type

Private s                                       As State


Public Sub Execute()
    Part01
    Part02
End Sub


Private Sub Part01()

    Initialise
    
    Dim myArea As Variant
    myArea = Array(0, 0, 999, 999)
    ReDim Preserve myArea(1 To 4)
    
    Set s.Display = BulbDisplay(myArea)
    s.Display.SwitchOff myArea
    
    Dim myItem As Variant
    Dim myItems As Iteritems: Set myItems = Iteritems(s.Instructions)
    Do
        Set myItem = myItems.curitem(0)
        Dim myS As seqC
        Set myS = myItem
    
        Select Case myS.First

            Case "on": s.Display.SwitchOn myS.Tail.Toarray

            Case "off":  s.Display.SwitchOff myS.Tail.Toarray

            Case "toggle": s.Display.Toggle myS.Tail.Toarray

        End Select

    Loop While myItems.MoveNext
    
    Dim myResult As Long: myResult = s.Display.LitBulbs(myArea)
    
    fmt.dbg "The answer for Day {0} Part 01 is 543903. Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult

End Sub


Private Sub Part02()

    Initialise
    
    Dim myArea As Variant
    myArea = Array(0, 0, 999, 999)
    ReDim Preserve myArea(1 To 4)
    
    Set s.Display = BulbDisplay(myArea)
    s.Display.SwitchOff myArea
    s.Display.UseBrightness = True
    
    Dim myItem As Variant
    Dim myItems As Iteritems: Set myItems = Iteritems(s.Instructions)
    Do
        Set myItem = myItems.curitem(0)
        Dim myS As seqC
        Set myS = myItem
    
        Select Case myS.First

            Case "on": s.Display.SwitchOn myS.Tail.Toarray

            Case "off":  s.Display.SwitchOff myS.Tail.Toarray

            Case "toggle": s.Display.Toggle myS.Tail.Toarray

        End Select

    Loop While myItems.MoveNext
    
    Dim myResult As Long: myResult = s.Display.Brightness(myArea)
  
    fmt.dbg "The answer for Day {0} Part 02 is 14687245. Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult


End Sub

Private Sub Initialise()

    Set s.Data = seqC(Filer.GetFileAsArrayOfStrings(AoCRawData & Year & TODAY))
   
    Set s.Instructions = s.Data _
        .MapIt(mpMultiReplace(Array(Array("through ", vbNullString), Array("turn ", vbNullString), Array(chars.twcomma, chars.twSpace), Array("  ", " ")))) _
        .MapIt(mpInner(mpsplit(chars.twSpace))) _
        .MapIt(mpInner(mpConvert(e_ConvertTo.m_Long)))
   
End Sub

', Array("turn ", vbNullString), Array(chars.twcomma, chars.twSpace), Array("  ", " ")))) _

