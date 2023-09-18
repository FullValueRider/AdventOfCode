VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Const SEPARATOR                     As String = "x"

Private Enum Edge

    IsLength = 0
    IsWidth = 1
    IsHeight = 2

End Enum

Public Enum FaceType

    LH = 0
    LW
    WH
    
End Enum


Private Type Face

    Area                            As Long
    Perimeter                       As Long
    
End Type


Private Type Properties

    Length                          As Long
    Width                           As Long
    Height                          As Long
    
    
End Type

Private p                           As Properties

Private Type State

    Faces(2&)                       As Face
    
End Type

Private s                           As State


Public Function Deb(ByVal ipDims As String) As Box
    With New Box
        Return .COnstructInstance(ipDims)
    End With
End Function


Public Function COnstructInstance(ByVal ipDims As String) As Box

    Dim myDimensions As Variant
    myDimensions = Split(ipDims, SEPARATOR)
    p.Length = CLng(myDimensions(Edge.IsHeight))
    p.Width = CLng(myDimensions(Edge.IsLength))
    p.Height = CLng(myDimensions(Edge.IsWidth))
    
    s.Faces(FaceType.LH).Area = p.Length * p.Height
    s.Faces(FaceType.LH).Perimeter = 2 * (p.Length + p.Height)
    
    s.Faces(FaceType.LW).Area = p.Length * p.Width
    s.Faces(FaceType.LW).Perimeter = 2 * (p.Length + p.Width)
    
    s.Faces(FaceType.WH).Area = p.Width * p.Height
    s.Faces(FaceType.WH).Perimeter = 2 * (p.Width + p.Height)
    
    Return Me
    
End Function


Public Function SurfaceArea() As Long
    SurfaceArea = (s.Faces(FaceType.LH).Area + s.Faces(FaceType.LW).Area + s.Faces(FaceType.WH).Area) * 2
End Function


Public Function Volume() As Long
    Volume = p.Length * p.Width * p.Height
End Function


Public Function AreaOfSmallestFace() As Long
    AreaOfSmallestFace = Min(s.Faces(FaceType.LH).Area, s.Faces(FaceType.LW).Area, s.Faces(FaceType.WH).Area)
End Function


Public Function PerimeterOfSmallestFace() As Long
    PerimeterOfSmallestFace = Min(s.Faces(FaceType.LH).Perimeter, s.Faces(FaceType.LW).Perimeter, s.Faces(FaceType.WH).Perimeter)
End Function


Private Function Min(ParamArray ipArgs() As Variant) As Long
    Dim myMin As Long
    myMin = ipArgs(0&)
    If ipArgs(1&) < myMin Then myMin = ipArgs(1&)
    If ipArgs(2&) < myMin Then myMin = ipArgs(2&)
    Min = myMin
End Function

Public Function WrappingSize() As Long

    WrappingSize = SurfaceArea + AreaOfSmallestFace
    
End Function


Public Function RibbonLength() As Long

    RibbonLength = Me.PerimeterOfSmallestFace + Me.Volume
    
End Function
