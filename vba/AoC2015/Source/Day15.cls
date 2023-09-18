VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Days")
Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C


Private Enum e_Field

    m_Cap = 1
    m_Dur
    m_fla
    m_text
    m_Cal
    m_Score
    m_Spoons
    


End Enum
Private Const TODAY             As String = "\Day14.txt"

Private Type State

    Data                            As seqC
    
    
End Type

Private s                           As State


Public Sub Execute()
    Part01
    Part02
End Sub

Private Sub Part01()

    Initialise
   
    
    Dim myResult As Long: myResult = 0
    fmt.dbg "The answer to Day {0} part 1 is {1}.  Found is 2640", VBA.Mid$(TODAY, 5, 2), myResult
    
End Sub

Private Sub Part02()

    Initialise
    
   
    Dim myResult As Long: myResult = 0
    fmt.dbg "The answer to Day {0} part 2 is 1102.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult

End Sub

Private Sub Initialise()

    Set s.Data = seqC(Filer.GetFileAsArrayOfStrings(AoCRawData & Year & TODAY)).MapIt(mpExecDeb(Reindeer)) _
        .MapIt(mpMultiReplace(Array(Array("capacity ", vbNullString), Array("durability ", vbNullString), Array("flavor ", vbNullString), Array("texture ", vbNullString), Array("calories ", vbNullString)))) _
        .MapIt(mpsplitToSeq(chars.twcomma)) _
        .MapIt(mpConvert(m_Long))
        
    'Set s.Ingredients = kvpc.Deb
  
    
    
End Sub
