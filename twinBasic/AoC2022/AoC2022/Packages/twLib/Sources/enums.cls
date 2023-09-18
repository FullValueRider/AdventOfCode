VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "enums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Provides a central location for Enums that have been wrapped in a Class"
Option Explicit

    
    
' Public Function Deb() As enums
'     With New enums
'         Set Deb = .constructInstance
'     End With
' End Function

' Public friend 


'@Description("Presets are 'constants' that are defined at runtime. This enumeration allows functions to be used in metthod signatures whose value is not known until runtime.")
Public Function Preset() As enumPreset
Attribute Preset.VB_Description = "Presets are 'constants' that are defined at runtime. This enumeration allows functions to be used in metthod signatures whose value is not known until runtime."
    Set Preset = enumPreset
End Function


'@Description("Standardized Texts that are used to inform the user of the internal state of the program.  Typical use is for errors")
Public Function Message() As enumMessage
Attribute Message.VB_Description = "Standardized Texts that are used to inform the user of the internal state of the program.  Typical use is for errors"
    Set Message = enumMessage
End Function

Public Function Types() As enumTypes
    Set Types = enumTypes
End Function

