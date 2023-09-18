VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A place for useful VBA things not explicitly provided by VBA"
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib")
   
'@ModuleDescription("A place for useful VBA things not explicitly provided by VBA")




'@Ignore ConstantNotUsed
' Public Function MaxLong() As Long:MaxLong = &H7FFFFFFF:End Function
' Public Function MinLong() As Long:MinLong = &HFFFFFFFF:End Function

' Public Function NotOkay() As Boolean: NotOkay = False: End Function
' Public Function Okay() As Boolean:Okay = True:End Function

' Public Function NotANumber() As String:NotANumber = "NaN":End Function

' Public Function AsOneItem(ByVal ipIterable As Variant) As Variant
'     AsOneItem = Array(ipIterable)
' End Function


Public Function IsTrue(ByVal ipexpression As Boolean) As Boolean
    IsTrue = ipexpression
End Function

Public Function IsNotTrue(ByVal ipexpression As Boolean) As Boolean
    IsNotTrue = Not IsTrue(ipexpression)
End Function

Public Function IsOkay() As Boolean
    IsOkay = True
End Function

' Yes isnotokay should be true
Public Function IsNotOkay() As Boolean
    IsNotOkay = True
End Function

Public Sub Swap(ByRef iopLHS As Variant, ByRef iopRHS As Variant)
    Dim myTmp As Variant
    If VBA.IsObject(iopLHS) Then
        Set myTmp = iopLHS
    Else
    	myTmp = iopLHS
    End If
    
    If VBA.IsObject(iopRHS) Then
        Set iopLHS = iopRHS
    Else
    	iopLHS = iopRHS
    End If
    If VBA.IsObject(myTmp) Then
        Set iopRHS = myTmp
    Else
    	iopRHS = myTmp
    End If
End Sub

'@Description(" = without the pain")
Public Sub Assign(ByRef opTo As Variant, ByRef ipFrom As Variant)
Attribute Assign.VB_Description = " = without the pain"
    
    If IsObject(ipFrom) Then
        
        Set opTo = ipFrom
        
        
    Else
        
        opTo = ipFrom
        
        
    End If
    
End Sub

Public Sub DoNothing()
    ' dummy call that does nothing
End Sub


'@Description("Takes a list of values or an Enumerable and returns a variant containing an Array")
' Public Function ToArray(ByRef ipCargo As Cargo) As Variant
' Attribute ToArray.VB_Description = "Takes a list of values or an Enumerable and returns a variant containing an Array"
'     ToArray = ToArray(CVar(ipArgArray))
' End Function

'Description("Performs the ToArray method but taking forwarded paramarray")
' Friend Function ToArray(ByRef ipVa As varArray) As varArray

'     Dim myVA As varArray
'     myVA.Cargo = Empty
'     ToArray = myVA
    
'     If BailOut.When(ArrayInfo.Ranks(ipVa.Cargo) = 0, ParamArrayIsNotAllocated) Then Exit Function
        
'     Dim myEnumerable As ArrayList
'     Set myEnumerable = ArrayList.Deb
    
'     Dim myIndex As Long
'     Dim myItem As Variant
    
'     Select Case True

'         Case ArrayInfo.HasItems(ipVa.Cargo)
'             '
'             ToArray = ipVa
'             Exit Function
        
'         Case VBA.IsArray(ipVa.Cargo(0))

'             myVA.Cargo = ipVa.Cargo(0)
'             ToArray = myVA
'             Exit Function


'         ' from this point onwards we are dealing with a param array containing 1 item that is not an array
'         Case TypeInfo.IsEnumerable(ipVa.Cargo(0), EnumerableType.AsItemByForEach)
            
'             For Each myItem In ipVa.Cargo(0)
                
'                 myEnumerable.Add myItem
                
                
'             Next
            
'             Dim myArray As Variant
'             ReDim myarray(0 To myEnumerable.Count - 1)
'             myVA.Cargo = myarray
'             myVA.Cargo = myEnumerable.Items
'             ToArray = myVA

            
'         Case TypeInfo.IsEnumerable(ipVa.Cargo(0), EnumerableType.AsItemByToArrayForEach)
        
'             myVA.Cargo = ipVa.Cargo(0).toarray
'             ToArray = myVA
'             Exit Function


'         Case TypeInfo.IsEnumerable(ipVa.Cargo(0), EnumerableType.AsKeyByForEach)
'             'ToDo: consider revising this to return an array of key/item pairs.
'             myVA.Cargo = ipVa.Cargo(0).Items
'             ToArray = myVA
'             Exit Function

        
'         Case TypeInfo.IsEnumerable(ipVa.Cargo(0), EnumerableType.AsItemByIndex)
        
'             myIndex = 0
    
'             On Error Resume Next
'             Do
                
'                 Sys.Assign myItem, ipVa.Cargo(0)(myIndex)
'                 myIndex = myIndex + 1
                
'             Loop Until Err.Number = 0
'             myEnumerable.Add myItem
'             Do While myIndex <= ipVa.Cargo(0).Count
            
               
'                 myIndex = myIndex + 1
'                 myEnumerable.Add ipVa.Cargo(0)(myIndex)
                
'             Loop
            
'             On Error GoTo 0
            
'             myVA.Cargo = myEnumerable.Items
           
'             ToArray = myVA
             
'         Case Else
'             ' Only a single item that is not enumerable in ipCargo.Items
'             ToArray = ipVa

            
'     End Select
    
'     Dim myArray As Variant
'     myIndex = 0
'     For Each myItem In myEnumerable
        
'         '@Ignore UnassignedVariableUsage
'         Sys.Assign myArray(myIndex), myItem
'         myIndex = myIndex + 1
        
        
'     Next
    
'     myVA.Cargo = myArray
'     ToArray = myVA
 
' End Function

'@Description("Takes a list of values or an Enumerable and returns a Collection)
' Public Function ToCollection(ByRef ipCargo As Cargo) As Collection
'     Set ToCollection = ToCollection(CVar(ipCargo))
' End Function

'@Ignore ParameterCanBeByVal
' Friend Function ToCollection(ByRef ipVA As varArray) As Collection
    
'     If BailOut.When(ArrayInfo.IsNotQueryable(ipVA.Cargo), ParamArrayIsNotAllocated) Then Exit Function
    
'     Dim myArray As varArray
'     myArray = ToArray(ipVA)
    
'     Dim myC As Collection
'     Set myC = New Collection
    
'     Dim myItem As Variant
'     For Each myItem In myArray.Cargo
'         myC.Add myItem
'     Next
    
'     Set ToCollection = myC
    
' End Function

'@Description("Takes a list of values or an Enumerable and returns an ArrayList")
' Public Function ToArrayList(ByRef ipPA As PA) As cArrayList
' Attribute ToArrayList.VB_Description = "Takes a list of values or an Enumerable and returns an ArrayList"
'     Set ToArrayList = ToArrayList(CVar(ipArgArray))
' End Function

' Friend Function ToArrayList(ByRef ipPA As PA) As ArrayList
    
'     If BailOut.When(ArrayInfo.IsNotQueryable(ipPA.Cargo), ParamArrayIsNotAllocated) Then Exit Function
    
'     Return ArrayList.Deb.AddRange(ipPA.Cargo)
    
   
    
' End Function

