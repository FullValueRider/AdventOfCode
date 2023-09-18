VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VentCoords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Day05")
    
    Private Enum CoordType
    
        ctX1 = 0
        ctY1
        ctX2
        ctY2
        
    End Enum
    
    Private Type State
        
        X1                              As Long
        Y1                              As Long
        X2                              As Long
        Y2                              As Long
        IsVentLineType As VentLineType
        
    End Type
    
    Private s As State
    
    
    
    Public Function Deb(ByVal ipCoords As String) As VentCoords
        With New VentCoords
            Set Deb = .ReadyToUseInstance(ipCoords)
        End With
    End Function
    
    
    Public Function ReadyToUseInstance(ByVal ipCoords As String) As VentCoords
        
        Dim mycoords As Variant = VBA.Split(ipCoords, Char.twComma)
        
        s.X1 = CLng(mycoords(ctX1))
        s.Y1 = CLng(mycoords(ctY1))
        s.X2 = CLng(mycoords(ctX2))
        s.Y2 = CLng(mycoords(ctY2))
            
        'Determine any equivalence between X1,X2 and Y1,Y2 so that we
        'can assign a type to the line
        
        Select Case True
            
            Case Is = ((s.X1 = s.X2) And (s.Y1 = s.Y2)):            s.IsVentLineType = VentLineType.vlPoint
            Case Is = (s.Y1 = s.Y2):                                s.IsVentLineType = VentLineType.vlHorizontal
            Case Is = (s.X1 = s.X2):                                s.IsVentLineType = VentLineType.vlVertical
            Case Else:                                              s.IsVentLineType = VentLineType.vlDiagonal
            
        End Select
        
        Set ReadyToUseInstance = Me
        
    End Function
    
    
    Public Function GetLineType() As VentLineType
        GetLineType = s.IsVentLineType
    End Function
    
    
    Public Function GetCoords() As Seq
        
        Dim myVentLine As Seq
        Set myVentLine = Seq.Deb
        
        Dim myXCoord As Long
        Dim myYCoord As Long
        Dim myYStep As Long
        myYStep = IIf(s.Y1 <= s.Y2, 1, -1)
        Dim myXStep As Long
        myXStep = IIf(s.X1 <= s.X2, 1, -1)
        
        Select Case s.IsVentLineType
            
            Case vlPoint
            
                myVentLine.Add Fmt.Txt("{0},{1}", s.X1, s.Y1)
                
                
            Case vlHorizontal
            
                For myXCoord = s.X1 To s.X2 Step myXStep
                    
                    myVentLine.Add Fmt.Txt("{0},{1}", myXCoord, s.Y1)
                    
                Next
            
                
            Case vlVertical
            
                For myYCoord = s.Y1 To s.Y2 Step myYStep
                    
                    myVentLine.Add Fmt.Txt("{0},{1}", s.X1, myYCoord)
                    
                Next
            
            
            Case vlDiagonal
            
                myYCoord = s.Y1
                For myXCoord = s.X1 To s.X2 Step myXStep
                    
                    myVentLine.Add Fmt.Txt("{0},{1}", myXCoord, myYCoord)
                    myYCoord = myYCoord + myYStep
                    
                Next
            
        End Select
        
        Set GetCoords = myVentLine
        
    End Function


    
   
