VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Day09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Days")
 Option Explicit
 '========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' This is a graph theory problem for finding the shortest distance between two points
 '
 ' https://www.freecodecamp.org/news/dijkstras-shortest-path-algorithm-visual-introduction/
 '
 ' The graph is non-directional
 ' The graph forbids using nodes more than once
 '

 Private Const TODAY                   As String = "\Day09.txt"

 Private Type State

     Data                               As seqC
     NodeLinks                          As KvpC
     Visited                            As KvpC
     Routes                             As seqC
     
 End Type

 Private s                              As State

 
 Public Sub Execute()
    Part01
    Part02
 End Sub

 Private Sub Part01()

     Initialise
        
     Set s.Routes = Permutations.ByKey(s.NodeLinks)
    
     Dim myMinDistance As Long
     myMinDistance = globals.MaxLong
     
     Dim myRoutes As Iteritems: Set myRoutes = Iteritems(s.Routes)
     Do
     
         DoEvents
         
         Dim myDistance As Long
         myDistance = 0
         
         Dim myroute As seqC
         Set myroute = myRoutes.curItem(0)
         
         Dim myNodes As Iteritems
         Set myNodes = Iteritems(myroute)
         Do While myNodes.MoveNext
             DoEvents
             myDistance = myDistance + s.NodeLinks.Item(myNodes.curItem(0)).Item(myNodes.curItem(-1))
         Loop
         
         If myDistance < myMinDistance Then
             myMinDistance = myDistance
         End If
    
     Loop While myRoutes.MoveNext
     
     Dim myResult As Long
     myResult = myMinDistance
     fmt.Dbg "The answer to Day {0} part 1 is 251.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult
     
 End Sub

 Private Sub Part02()

     Initialise
     Set s.Routes = Permutations.ByKey(s.NodeLinks)
    
     Dim myMaxDistance As Long
     myMaxDistance = globals.MinLong
     
     Dim myRoutes As Iteritems
     Set myRoutes = Iteritems(s.Routes)
     Do
     
         DoEvents
         
         Dim myDistance As Long
         myDistance = 0
         
         Dim myroute As seqC
         Set myroute = myRoutes.curItem(0)
         
         Dim myNodes As Iteritems
         Set myNodes = Iteritems(myroute)
         Do While myNodes.MoveNext
             DoEvents
             myDistance = myDistance + s.NodeLinks.Item(myNodes.curItem(0)).Item(myNodes.curItem(-1))
         Loop
         
         If myDistance > myMaxDistance Then
             myMaxDistance = myDistance
         End If
     
     Loop While myRoutes.MoveNext
    
     Dim myResult As Long
     myResult = myMaxDistance
             
     fmt.Dbg "The answer to Day {0} part 2 is 898.  Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult

 End Sub

 Private Sub Initialise()
 
     Set s.Data = seqC(Filer.GetFileAsArrayOfStrings(AoCRawData & Year & TODAY)) _
         .mapIt(mpMultiReplace(Array(Array(" = ", chars.twcomma), Array(" to ", chars.twcomma)))) _
         .mapIt(mpsplit(chars.twcomma))
     
     Set s.Visited = KvpC.Deb
     Set s.NodeLinks = BuildNodeToNodeDictionary(s.Data)
     
 End Sub
 
 Public Function BuildNodeToNodeDictionary(ByRef ipLinks As seqC) As KvpC
 
     ' Builds a kvpc of kvpc
     ' The outer Kvpc us start node vs Kvpc
     ' The inner Kvpc is destination vs distance
     
     Dim myStartNodes As KvpC
     Set myStartNodes = KvpC.Deb
     
     Dim myLinks As Iteritems
     Set myLinks = Iteritems.Deb(ipLinks)
     
     Do
     
         Dim myLink As seqC
         Set myLink = myLinks.curItem(0)
         
         Dim myStartNode As String
         myStartNode = myLink.Item(1)
         
         Dim myEndNode As String
         myEndNode = myLink.Item(2)
         
         Dim myDistance As Long
         myDistance = myLink.Item(3)
         
         If myStartNodes.HoldsKey(myStartNode) Then
             myStartNodes.Item(myStartNode).Add myEndNode, myDistance
         Else
             myStartNodes.Add myStartNode, KvpC.Deb.Add(myEndNode, myDistance)
         End If
         
         If myStartNodes.HoldsKey(myEndNode) Then
             myStartNodes.Item(myEndNode).Add myStartNode, myDistance
         Else
             myStartNodes.Add myEndNode, KvpC.Deb.Add(myStartNode, myDistance)
         End If
         
     Loop While myLinks.MoveNext
     
     Set s.Routes = seqC.Deb
    
     Set BuildNodeToNodeDictionary = myStartNodes
     
 End Function
 



