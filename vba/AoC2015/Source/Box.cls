VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Classes")
Option Explicit

Const SEPARATOR                     As String = "x"
Const MY_TYPENAME                   As String = "Box"

Private Type BoxData
    Edges                           As seqC
    Surfaces                        As seqC
    Perims                          As seqC
End Type
     
Private Type State
    BoxInfo                         As BoxData
End Type

Private s                           As State

'@DefaultMember
Public Function Deb(ByVal ipDims As String) As Box
Attribute Deb.VB_UserMemId = 0
    With New Box
        Set Deb = .ConstructInstance(ipDims)
    End With
End Function


Public Function ConstructInstance(ByVal ipDims As String) As Box

    With s.BoxInfo
        Set .Edges = seqC(VBA.Split(ipDims, SEPARATOR)).MapIt(mpConvert(e_ConvertTo.m_Long))
        Set .Surfaces = seqC(.Edges.Item(1) * .Edges.Item(2), .Edges.Item(2) * .Edges.Item(3), .Edges.Item(3) * .Edges.Item(1))
        Set .Perims = seqC(2 * (.Edges.Item(1) + .Edges.Item(2)), 2 * (.Edges.Item(2) + .Edges.Item(3)), 2 * (.Edges.Item(3) + .Edges.Item(1)))
    End With
    
    
    Set ConstructInstance = Me
    
End Function


Public Function SurfaceArea() As Long
    SurfaceArea = s.BoxInfo.Surfaces.ReduceIt(rdSum) * 2
End Function


Public Function Volume() As Long
    Volume = s.BoxInfo.Edges.Item(1) * s.BoxInfo.Edges.Item(2) * s.BoxInfo.Edges.Item(3)
End Function


Public Function AreaOfSmallestFace() As Long
    AreaOfSmallestFace = s.BoxInfo.Surfaces.ReduceIt(rdMinNum.Deb)
End Function


Public Function SmallestPerimeter() As Long
    SmallestPerimeter = s.BoxInfo.Perims.ReduceIt(rdMinNum.Deb)
End Function


Public Function WrappingSize() As Long
    WrappingSize = SurfaceArea + AreaOfSmallestFace
End Function


Public Function RibbonLength() As Long
    RibbonLength = SmallestPerimeter + Volume
End Function


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property

