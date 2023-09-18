VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type State

    Ws                                      As Excel.Worksheet
    OrgX                                    As Long
    OrgY                                    As Long
End Type

Private s                                   As State


Private Sub Class_Initialize()

    Dim myapp As Excel.Application
    Set myapp = New Excel.Application
    myapp.Visible = True
    Dim mywb As Excel.Workbook
    Set mywb = myapp.Workbooks.Add(xlWBATWorksheet)
    Set s.Ws = mywb.ActiveSheet
    myapp.Visible = True
    s.OrgX = 20
    s.OrgY = 20
    
End Sub

Private Sub Class_Terminate()

    Dim mywb As Excel.Workbook
    Dim myapp As Excel.Application
    
    Set mywb = s.Ws.Parent
    Set myapp = s.Ws.Application
    
    Set s.Ws = Nothing
    mywb.Close
    Set mywb = Nothing
    myapp.Quit
    Set myapp = Nothing
    
End Sub



Public Sub PlaceTile(ByVal ipX As Long, ByVal ipY As Long, ByVal ipPiece As String)

    Dim myX As Long
    Do While ipX + s.OrgX < 1
        ExpandAtLeft
    Loop
    myX = ipX + s.OrgX
    
    Dim myY As Long
    Do While (ipY * -1) + 20 < 1
        ExpandAtTop
    Loop
    myY = ipY + s.OrgY
    
    ' Makes 20,20 equivalent to 0,0
    s.Ws.Cells.Item(myY, myX).Value = ipPiece
    
End Sub

Public Sub ExpandAtLeft()

    s.Ws.Columns.Item("A:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
End Sub


Public Sub ExpandAtTop()


    s.Ws.Rows.Item("1:5").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    'Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub
