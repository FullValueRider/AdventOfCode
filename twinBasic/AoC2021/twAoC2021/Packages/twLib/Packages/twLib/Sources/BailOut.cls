VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BailOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Admin")

' The bailout class combines guard statments and error logging.

'@Ignore EmptyMethod
Private Sub Class_Initialize()
End Sub

'@Description("Sets message and global.res.status to not okay if ipBailout is true. Returns ipBailout ")
Public Function When _
( _
    ByVal ipBailOut As Boolean, _
    Optional ByVal ipMsgOrd As MsgOrd = Unspecified, _
    Optional ByVal ipAltMessage As String, _
    Optional ByVal ipMsgArgsLyst As Variant _
) As Boolean
Attribute When.VB_Description = "Sets message and global.res.status to not okay if ipBailout is true. Returns ipBailout "
    
    If ipBailOut Then
       
        '@Ignore FunctionReturnValueDiscarded
        'Res.SetNotOkay ipMsgOrd, ipAltMessage, ipMsgArgsLyst
        
    End If

    'ToDo: use a msgbox with report in a scrollable box.
    If BAILOUT_RAISES_ERRORS Then
        
        Err.Raise _
            vbObjectError + 17, _
            "Lib Error", _
            "ToDo: add code to produce the report, meanwhile have a look at the result variable" 'Res.Report
        Dim myResult As Result = Result ' allows res tto be seen in the debugger at this point
    End If

    When = ipBailOut
    
End Function

' '@Description("Sets message and global.res.status to okay. ")
' Public Function CannotProceed _
' ( _
'     ByVal ipCannotProceed As Boolean, _
'     Optional ByVal ipMsgOrd As MsgOrd = Unspecified, _
'     Optional ByVal ipMsgArgsLyst As Variant, _
'     Optional ByVal ipAltMessage As String _
' ) As Boolean
' Attribute CannotProceed.VB_Description = "Sets message and global.res.status to okay. Returns ipNotify."
        
'     If ipCannotProceed Then
        
'         Res.SetOkay ipMsgOrd, ipAltMessage, ipMsgArgsLyst
'         '@Ignore FunctionReturnValueDiscarded
        
        
'     End If
    
'     CannotProceed = ipCannotProceed
    
' End Function


' Friend Sub SetGlobalResMsg _
' ( _
'     ByVal ipMsgOrd As MsgOrd, _
'     Optional ByVal ipArgLyst As Variant, _
'     Optional ByVal ipAltMessage As String _
' )
    
'     Dim myMessage As String
'     If VBA.Len(ipAltMessage) = 0 Then
        
'         myMessage = enums.Message.Value(ipMsgOrd)
            
        
'     Else
            
'         myMessage = ipAltMessage
            
        
'     End If
    
'     Dim myErrorText As String
'     Dim myargs As Variant
'     If IsMissing(ipArgLyst) Then
        
'         myErrorText = myMessage
        
'     Else
        
'         myErrorText = Fmt.TxtByFPA(myMessage, oArr(ipArgLyst))
        
'     End If
    
'     Res.SetNotOkay ipmessage:=myErrorText
    
' End Sub
