VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Journey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
Option Explicit

Private Const NORTH                 As String = "^"
Private Const SOUTH                 As String = "v"
Private Const WEST                  As String = "<"
Private Const EAST                  As String = ">"


Private Type State
    x                               As Long
    y                               As Long
    Directions                      As KvpC ' x and y increments for moving n,s,e or w
End Type

'@Ignore MoveFieldCloserToUsage
Private s                           As State


Private Type Properties
    Path                            As seqHC ' record of the sequence of x,y coords visited
    Visits                          As kvpha ' record of how many times each x,y coord was visited
End Type

Private p                           As Properties

'@DefaultMember
Public Function Deb(ByVal ipInstructions As SeqA) As Journey
Attribute Deb.VB_UserMemId = 0
    With New Journey
        Set Deb = .ConstructInstance(ipInstructions)
    End With
End Function


Public Function ConstructInstance(ByVal ipMoves As SeqA) As Journey

    s.x = 0
    s.y = 0
    
    Set s.Directions = KvpC.Deb
    s.Directions.Add NORTH, Array(0, 1)
    s.Directions.Add SOUTH, Array(0, -1)
    s.Directions.Add EAST, Array(-1, 0)
    s.Directions.Add WEST, Array(1, 0)
    
    Dim myMoves As Iteritems: Set myMoves = Iteritems(ipMoves)
    Do
        Dim myMove As String: myMove = myMoves.curitem(0)
        s.x = s.x + s.Directions.Item(myMove)(0)
        s.y = s.y + s.Directions.Item(myMove)(1)
        
        Dim myCoords As String
        myCoords = fmt.Text("{0},{1}", s.x, s.y)
        
        UpdatePath myCoords
        UpdateVisits myCoords
        
    Loop While myMoves.MoveNext

    Set ConstructInstance = Me

End Function


Public Sub UpdatePath(ByVal ipCoords As String)

    If p.Path Is Nothing Then
    
        Set p.Path = seqHC.Deb
       
        p.Path.Add "0,0"
    
    End If

    p.Path.Add ipCoords
           
End Sub

Public Sub UpdateVisits(ByVal ipCoords As String)

    ' Track is a Kvp of Key:Coordinates as string vs Value: Kvp of Instruction
    ' count when coordinates were visited
    ' i.e. When a house is visited more than once we know at what instruction
    ' steps the visits were
    

    If p.Visits Is Nothing Then
        Set p.Visits = kvpha.Deb
    End If
    
    
    If Not p.Visits.Exists(ipCoords) Then
        p.Visits.Add ipCoords, 1
    Else
        p.Visits.Item(ipCoords) = p.Visits.Item(ipCoords) + 1
    End If
    
End Sub

'@Description("Returns a KvpC of coordinates vs number of times the coordinate was visited")
Public Property Get Visits() As kvpha
Attribute Visits.VB_Description = "Returns a KvpC of coordinates vs number of times the coordinate was visited"
    Set Visits = p.Visits
End Property

'@Description("Returns a seq of coordinates in th order in which they were visited")
Public Property Get Path() As seqHC
Attribute Path.VB_Description = "Returns a seq of coordinates in th order in which they were visited"
    Set Path = p.Path
End Property
