VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventConsumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
'@Ignore EncapsulatePublicField
Public WithEvents myProvider As EventProvider
Attribute myProvider.VB_VarHelpID = -1


Private Sub Class_Initialize()

    Set myProvider = New EventProvider
    
End Sub

Private Sub myProvider_Success(ByVal ipMessage As String)

    Debug.Print "MyProvider sent the message" & ipMessage

End Sub
