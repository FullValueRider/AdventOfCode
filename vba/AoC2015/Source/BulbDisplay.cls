VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BulbDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
Option Explicit

Private Enum Area

    OrgX = 1
    OrgY
    MaxX
    MaxY
    
End Enum


Private Type State

    Array           As Variant
    UseBrightness   As Boolean
    
End Type

Private s As State

'@defaultmember
Public Function Deb(ByVal ipArea As Variant, Optional ByVal ipUseBrightness As Boolean = False) As BulbDisplay
Attribute Deb.VB_UserMemId = 0
    With New BulbDisplay
        Set Deb = .ConstructInstance(ipArea, ipUseBrightness)
    End With
End Function


Friend Function ConstructInstance(ByVal ipArea As Variant, ByVal ipUseBrightness As Boolean) As BulbDisplay

    ReDim s.Array(ipArea(Area.OrgX) To ipArea(Area.MaxX), ipArea(Area.OrgY) To ipArea(Area.MaxY))
    s.UseBrightness = ipUseBrightness
    Set ConstructInstance = Me
    
End Function


Public Property Let UseBrightness(ByVal ipValue As Boolean)
    s.UseBrightness = ipValue
End Property

Public Property Get UseBrightness() As Boolean
    UseBrightness = s.UseBrightness
End Property


Public Sub SwitchOn(ByVal ipArea As Variant)
    ' iparea is a variant containg array(0 to 3) of x1,y1 x2,y2
    Dim myX As Long
    For myX = ipArea(Area.OrgX) To ipArea(Area.MaxX)
    
        Dim myY As Long
        For myY = ipArea(Area.OrgY) To ipArea(Area.MaxY)
        
            If s.UseBrightness Then
            
                s.Array(myX, myY) = s.Array(myX, myY) + 1
                
            Else
            
                s.Array(myX, myY) = 1
                
            End If
            
        Next
        
    Next
    
End Sub


Public Sub SwitchOff(ByVal ipArea As Variant)
' iparea is a variant containg array(0 to 3) of x1, y1 x2,y2
    Dim myX As Long
    For myX = ipArea(Area.OrgX) To ipArea(Area.MaxX)
    
        Dim myY As Long
        For myY = ipArea(Area.OrgY) To ipArea(Area.MaxY)
        
            If s.UseBrightness Then
            
                s.Array(myX, myY) = s.Array(myX, myY) - 1
                If s.Array(myX, myY) < 0 Then s.Array(myX, myY) = 0
                
            Else
            
                s.Array(myX, myY) = 0
                
            End If
            
        Next
        
    Next
    
End Sub


Public Sub Toggle(ByVal ipArea As Variant)

' iparea is a variant containg array(0 to 3) of x1, y1 x2,y2
    Dim myX As Long
    For myX = ipArea(Area.OrgX) To ipArea(Area.MaxX)
    
        Dim myY As Long
        For myY = ipArea(Area.OrgY) To ipArea(Area.MaxY)
        
            If s.UseBrightness Then
            
                s.Array(myX, myY) = s.Array(myX, myY) + 2
                
            Else
            
                s.Array(myX, myY) = IIf(s.Array(myX, myY) = 0, 1, 0)
                
            End If
            
        Next
        
    Next
    
End Sub


Public Function LitBulbs(ByVal ipArea As Variant) As Long

    Dim myX As Long
    For myX = ipArea(Area.OrgX) To ipArea(Area.MaxX)
    
        Dim myY As Long
        For myY = ipArea(Area.OrgY) To ipArea(Area.MaxY)
        
            Dim myCount As Long
            If s.Array(myX, myY) > 0 Then myCount = myCount + 1
            
        Next
        
    Next
    
    LitBulbs = myCount

End Function


Public Function Brightness(ByVal ipArea As Variant) As Long

    Dim myX As Long
    For myX = ipArea(Area.OrgX) To ipArea(Area.MaxX)
    
        Dim myY As Long
        For myY = ipArea(Area.OrgY) To ipArea(Area.MaxY)
        
            Dim myCount As Long
            myCount = myCount + s.Array(myX, myY)
            
        Next
        
    Next
    
    Brightness = myCount

End Function
