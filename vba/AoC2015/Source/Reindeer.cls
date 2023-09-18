VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Reindeer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Helpers")
Option Explicit

Private Const FLYING                            As Boolean = True
Private Const NOT_FLYING                        As Boolean = False
Private Const MY_TYPENAME                       As String = "Reindeer"

Public Enum e_ReindeerInfo

    m_name = 0
    m_FlightSpeed
    m_FlightTime
    m_RestTime

End Enum
    

Private Type State

    Activity                    As Boolean
    Name                        As String
    FlightSpeed                 As Long
    FlightTime                  As Long
    RestTime                    As Long
    
End Type

Private s                       As State

'@DefaultMember
Public Function Deb(ByRef ipInfo As String) As Reindeer
Attribute Deb.VB_UserMemId = 0
    With New Reindeer
        Set Deb = .ConstructInstance(ipInfo)
    End With
End Function

Public Function ConstructInstance(ByRef ipInfo As String) As Reindeer

    Dim myInfo As String
    myInfo = strs.MultiReplacer(ipInfo, Array(Array(" can fly", vbNullString), Array(" km/s for", vbNullString), Array(" seconds, but then must rest for", vbNullString), Array(" seconds.", vbNullString)))
    Dim myTmp As Variant: myTmp = VBA.Split(myInfo, chars.twSpace)
       
    s.Name = myTmp(e_ReindeerInfo.m_name)
    s.FlightSpeed = VBA.CLng(myTmp(e_ReindeerInfo.m_FlightSpeed))
    s.FlightTime = VBA.CLng(myTmp(e_ReindeerInfo.m_FlightTime))
    s.RestTime = VBA.CLng(myTmp(e_ReindeerInfo.m_RestTime))
    s.Activity = FLYING
    
    Set ConstructInstance = Me
    
End Function

Public Function Distance(ByRef ipTime As Long) As Long
    
    Dim myResult As Long: myResult = (ipTime \ (s.FlightTime + s.RestTime)) * s.FlightSpeed * s.FlightTime
    
    Dim myRemainder As Long: myRemainder = ipTime Mod (s.FlightTime + s.RestTime)
    
    If myRemainder < s.FlightTime Then
    
        myResult = myResult + myRemainder * s.FlightSpeed
    
    Else
        
        myResult = myResult + s.FlightTime * s.FlightSpeed
        
    End If
    
    Distance = myResult
    
End Function

Public Property Get Name() As String
    Name = s.Name
End Property

Public Property Get FlightSpeed() As Long
    FlightSpeed = s.flighSpeed
End Property

Public Property Get FlightTime() As Long
    FlightTime = s.FlightTime
End Property


Public Property Get RestTime() As Long
    RestTime = s.RestTime
End Property

Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property
