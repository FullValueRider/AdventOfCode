VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "enumPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Const MY_TYPENAME                   As String = "enumPreset"

Public Enum PresetOrd
    
    NaN = 1
    MaxLong
    MinLong
    MaxLongLong
    MinLongLong
    
    ' Array related
     ArrayFirstRank
    
    ' Result related
    ResultStatusOkay
    ResultStatusNotOkay

    'Kvp related
    KvpInsertIndexIsAdd
    DefaultDec
    DefaultInc
    DefaultAdjust
    
End Enum

Private Type State
    
    Values                      As Hkvp' values may be executable objects
    Names                       As Hkvp
    
End Type

Private s                       As State

Private Sub Class_Initialize()
        
    If Not Me Is enumPreset Then
        
        Err.Raise _
            enums.Message.Value(vbCannotPerformRequestedOperation), _
            "Static Class", _
            "Use the Predeclared instance only"
            
        Exit Sub
        
    End If
    
    SetupDictionaries
        
End Sub


'@Description("Returns the PredeclaredId")
Public Function Deb() As enumPreset
Attribute Deb.VB_Description = "Returns the PredeclaredId"
    If s.Values Is Nothing Then SetupDictionaries
    Set Deb = Me
End Function


Private Sub SetupDictionaries()
  
    Set s.Values = Hkvp.Deb
    With s.Values
        
        'Number related
        .Add NaN, "NaN"
        .Add MaxLong, &H7FFFFFFF&
        .Add MinLong, &H80000000&
        .Add MaxLongLong, &H7FFFFFFFFFFFFFFF^
        .Add MinLongLong, &H8000000000000000^
        ' Array related
        .Add ArrayFirstRank, 1
        
        ' Result related
        .Add ResultStatusOkay, True
        .Add ResultStatusNotOkay, False

        'Kvp related
        .Add KvpInsertIndexIsAdd, 0
        .Add DefaultDec, 1
        .Add DefaultInc, 1
        .Add DefaultAdjust, 1
        
    End With
    
    Set s.Names = Hkvp.Deb
    With s.Names
    ' Number related
        .Add NaN, "NaN"
        .Add MaxLong, "MaxLOng"
        .Add MinLong, "MinLOng"
        .Add MaxLongLong, "MaxLongLong"
        .Add MinLongLong, "MinLongLong"
        
        ' Attay related
       .Add ArrayFirstRank, "ArrayFirstRank"
        
        ' Result related
        .Add ResultStatusOkay, "ResultStatusOkay"
        .Add ResultStatusNotOkay, "ResultStatusNotOkay"

        'Kvp related
        '.Add KvpInsertIndexIsAdd, "KvpInsertIndexIsAdd"
        '.Add DefaultDec, "DefaultDec"
        '.Add DefaultInc, "DefaultInc"
        '.Add DefaultAdjust, "DefaultAdjust"
        
    End With
            
        
End Sub


Private Sub ValidateOrdinal(ByVal ipPresetOrd As PresetOrd)
   
    If s.Values.LacksKey(ipPresetOrd) Then
        
         Err.Raise _
            enums.Message.Value(vbInvalidOrdinal), _
            "No such member", _
            "The enumeration member does not exist"
            
        Exit Sub
    End If
    
End Sub


Public Function Values() As Variant
    Return s.Values.Items
End Function


' Public Function Members() As Variant
'     Members = s.Values.Keys
' End Function


Public Function Names() As Variant
    Return s.Names.Items
End Function


Public Sub AddValue(ByVal ipPresetOrd As PresetOrd, ByVal ipPreset As Variant)
    ValidateOrdinal ipPresetOrd
    Set s.Values.Item(ipPresetOrd) = ipPreset
End Sub

Public Function Name(ByVal ipPresetOrd As PresetOrd) As String
    ValidateOrdinal ipPresetOrd
    Return s.Names.Item(ipPresetOrd)
End Function

Public Function Ord(ByVal ipPresetOrd As Variant) As PresetOrd
    
    
    Dim myIndex As Long
    myIndex = 1
    Dim myItem As Variant
    For Each myItem In s.Values.Items
        
        If myItem = ipPresetOrd Then
            
            Return myIndex

            
        End If
        
        myIndex = myIndex + 1
        
    Next
    
    Return -1

End Function


'@DefaultMember'
Public Property Get Value(ByVal ipPresetOrd As PresetOrd) As Variant
    
    ValidateOrdinal ipPresetOrd
    If IsObject(s.Values.Item(ipPresetOrd)) Then
        
       Return s.Values.Item(ipPresetOrd)
        
        
    Else
       
        Return s.Values.Item(ipPresetOrd)
        
        
    End If
    
End Property


Public Function HoldsMember(ByVal ipPresetOrd As Long) As Boolean
    Return s.Names.HoldsItem(ipPresetOrd)
End Function


Public Function LacksMember(ByVal ipPresetOrd As Long) As Boolean
    Return Not s.Names.LacksItem(ipPresetOrd)
End Function


Public Property Get TypeName() As String
    Return MY_TYPENAME
End Property


Public Property Get Count() As Long
    Return s.Names.Count
End Property
