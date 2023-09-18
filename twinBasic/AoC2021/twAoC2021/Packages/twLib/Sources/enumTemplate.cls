VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "enumTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@IgnoreModule
'@PredeclaredId
'@Exposed
'===========================================================================================================
' This templatesimplifies the setup of a wrapped enumeration
' 1. find and replace XXXXXXXX with the enumeration name
' 2. In the enum definition    - find and replace memberx with your name
' 3. In Method SetupDictionaries - assign the values to each enumeration memeber.
' In Method SetupDictionaries define the value assigned to each enum member
'============================================================================================================

Const EnumerationName                           As String = "enumTemplate"

' The ordinal values should start at 1
Public Enum TemplateOrd
    sUnspecified = 1
    Member2
    Member3
    Member4
    Member5
    Member6
    Member7
    Member8
    Member9
    Member10
    Member11
    Member12
    Member13
    Member14
    Member15
    Member16
    Member17
    Member18
    Member19
    Member20
    Member21
    Member22
    Member23
    Member24
    Member25
    
End Enum

Private Type State
        
    Values                  As Hkvp' Values may be executable objects
    Names                   As Hkvp
        
End Type

Private s                   As State

Private Sub Class_Initialize()
        
    If Not Me Is enumTemplate Then
        
        Err.Raise _
            enums.Message.Value(vbCannotPerformRequestedOperation), _
            "Static Class", _
            "Use the Predeclared instance only"
            
        Exit Sub
        
    End If
    
    SetupDictionaries
        
End Sub

'@Description("Returns the PredeclaredId")
Public Function Deb() As enumTemplate
    If s.Values Is Nothing Then SetupDictionaries
    Set Deb = Me
End Function

Private Sub SetupDictionaries()
        
    ' Values associated with the enumeration member may be defined at runtime
    Set s.Values = New Hkvp
    With s.Values
        
    .Add sUnspecified, Empty ' or value
    .Add Member2, Empty ' or value
    .Add Member3, Empty ' or value
    .Add Member3, Empty ' or value
    .Add Member4, Empty ' or value
    .Add Member5, Empty ' or value
    .Add Member6, Empty ' or value
    .Add Member7, Empty ' or value
    .Add Member8, Empty ' or value
    .Add Member9, Empty ' or value
    .Add Member10, Empty ' or value
    .Add Member11, Empty ' or value
    .Add Member12, Empty ' or value
    .Add Member13, Empty ' or value
    .Add Member14, Empty ' or value
    .Add Member15, Empty ' or value
    .Add Member16, Empty ' or value
    .Add Member17, Empty ' or value
    .Add Member18, Empty ' or value
    .Add Member19, Empty ' or value
    .Add Member20, Empty ' or value
    .Add Member21, Empty ' or value
    .Add Member22, Empty ' or value
    .Add Member23, Empty ' or value
    .Add Member24, Empty ' or value
    .Add Member25, Empty ' or value
        
    End With
        
    Set s.Names = New Hkvp
    With s.Names
        
    .Add sUnspecified, "Member1"
    .Add Member2, "Member2"
    .Add Member3, "Member3"
    .Add Member4, "Member4"
    .Add Member5, "Member5"
    .Add Member6, "Member6"
    .Add Member7, "Member7"
    .Add Member8, "Member8"
    .Add Member9, "Member9"
    .Add Member10, "Member10"
    .Add Member11, "Member11"
    .Add Member12, "Member12"
    .Add Member13, "Member13"
    .Add Member14, "Member14"
    .Add Member15, "Member15"
    .Add Member16, "Member16"
    .Add Member17, "Member17"
    .Add Member18, "Member18"
    .Add Member19, "Member19"
    .Add Member20, "Member20"
    .Add Member21, "Member21"
    .Add Member22, "Member22"
    .Add Member23, "Member23"
    .Add Member24, "Member24"
    .Add Member25, "Member25"
        
    End With
        
End Sub


Private Sub ValidateOrdinal(ByVal ipTemplateOrd As TemplateOrd)
   
    If s.Values.LacksKey(ipTemplateOrd) Then
        
         Err.Raise _
            enums.Message.Value(vbInvalidOrdinal), _
            "No such member", _
            "The enumeration member does not exist"
            
        Exit Sub
    End If
    
End Sub


Public Function Names() As Variant
    Names = s.Names.Items
End Function


Public Function Members() As Variant
    Members = s.Names.Keys
End Function


Public Function Values() As Variant
    Values = s.Values.Items
End Function


Public Sub AddValue(ByVal ipTemplateOrd As TemplateOrd, ByVal ipTemplate As Variant)
    ValidateOrdinal ipTemplateOrd
    Set s.Values.Item(ipTemplateOrd) = ipTemplate
End Sub


Public Property Get Ord(ByVal ipValue As Variant) As TemplateOrd
    
    Ord = 0
    Dim myIndex As Long
    myIndex = 1
    Dim myItem As Variant
    For Each myItem In s.Values.Items
        
        If myItem = ipValue Then
            
            Ord = myIndex
            Exit Property
            
        End If
        
        myIndex = myIndex + 1
        
    Next
    
End Property


Public Function Name(ByVal ipTemplateOrd As TemplateOrd) As String
    ValidateOrdinal ipTemplateOrd
    Name = s.Names.Item(ipTemplateOrd)
End Function


'@DefaultMember
Public Property Get Value(ByVal ipTemplateOrd As TemplateOrd) As Variant
    
    ValidateOrdinal ipTemplateOrd
    If IsObject(s.Values) Then
        
        Set Value = s.Values.Item(ipTemplateOrd)
        
        
    Else
       
        Value = s.Values.Item(ipTemplateOrd)
        
        
    End If
    
End Property


Public Property Get HoldsMember(ByVal ipTemplateOrd As TemplateOrd) As Boolean
    HoldsMember = s.Names.HoldsKey(ipTemplateOrd)
End Property


Public Property Get LacksMember(ByVal ipTemplateOrd As TemplateOrd) As Boolean
    LacksMember = s.Names.LacksKey(ipTemplateOrd)
End Property


Public Property Get EnumName() As String
    EnumName = EnumerationName
End Property


Public Property Get Count() As Long
    Count = s.Names.Count
End Property
