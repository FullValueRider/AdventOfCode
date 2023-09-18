VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Day02")


Private Const InputData     As String = "Day02.txt"

Private Type State

    Data                        As Seq
    
End Type

Private s                       As State


Public Sub Execute()
    
    Part01
    Part02
    
End Sub

Public Sub Part01()
       
    Set s.Data = Seq.Deb _
        .AddRange(Filer.GetFileAsArrayOfStrings(AoC2021Data & InputData, vbCrLf)) _
        .MapIt(mpStrConv(vbLowerCase)) _
        .MapIt(mpSplitToSubStr(ToSeq, Char.twSpace))
    
    Dim myItem As Variant
    Dim myX As Long
    Dim myY As Long
    
    For Each myItem In s.Data.Items
       
       Dim myMove As Long = CLng(myItem(2))
       
       Select Case myItem(1)
           
           Case "forward":              myX += myMove
           Case "up":                   myY -= myMove
           Case "down":                 myY += myMove
               
               
       End Select
       
    Next
    
    Dim myResult As Long = myX * myY
    
    Fmt.Dbg "The answer to Day 02 part 1 is 1507611.  Found is {0}", myResult
    
End Sub

Public Sub Part02()
    
    Set s.Data = _
        Seq _
            .Deb _
            .AddRange(Filer.GetFileAsArrayOfStrings(AoC2021Data & InputData, vbCrLf)) _
            .MapIt(mpStrConv(vbLowerCase)) _
            .MapIt(mpDeDup) _
            .MapIt(mpSplitToSubStr(ToSeq, Char.twSpace))
    
    
    Dim myX As Long
    Dim myY As Long
    Dim myAim As Long
    Dim myI As IterItems = IterItems(s.Data)
    Do
        Dim mymove As Long = CLng(myI.Item()(2))
        
        Select Case myI.Item()(1)
            
            Case "forward"
            
                myX += mymove
                myY += (myAim * mymove)
            
                
            Case "up":              myAim -= mymove
            Case "down":            myAim += mymove
                
                
        End Select
    
    Loop While myI.MoveNext
        
    Dim myResult As Long = myX * myY
     
    Fmt.Dbg "The answer to Day 02 part 2 is 1880593125. Found Is {0}", myResult
    
End Sub
