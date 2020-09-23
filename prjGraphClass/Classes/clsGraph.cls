VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Size
    cx As Long
    cy As Long
End Type
Private Type RECT
    Left As Double
    Top As Double
    Right As Double
    Bottom As Double
End Type

Private TextSz As Size
Private Declare Function GetTextExtentPoint32 Lib "gdi32" _
                  Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                                 ByVal lpsz As String, _
                                                 ByVal cbString As Long, _
                                                 lpSize As Size) As Long

Private cObject As PictureBox
Private cFileName As String
Private cXMin As Double
Private cYMin As Double
Private cXMax As Double
Private cYMax As Double
Private cXVals() As Double
Private cYVals() As Double
Private numVal As Long
Private k As Long, SclX, SclY

Private cXAxisLarge As Single
Private cXAxisSmall As Single
Private cYAxisLarge As Single
Private cYAxisSmall As Single
Private cXAxisNumber As Single
Private cYAxisNumber As Single

Private cXAxis As clsAxis
Private cYAxis As clsAxis

Private TwpX As Long
Private TwpY As Long
Private ScaleArea As RECT
Private Const DotColor As Long = 5263440
Private Const AxisColor As Long = 8553090

Public Property Let GraphObject(objGraph As PictureBox)
    Set cObject = objGraph
    cObject.Cls
    cObject.Picture = Nothing
End Property

Public Property Let FileName(cVal As String)
    cFileName = cVal
End Property

Public Property Get XAxis() As clsAxis
    Set XAxis = cXAxis
End Property

Public Property Get YAxis() As clsAxis
    Set YAxis = cYAxis
End Property

Public Property Get XRatio() As Double
    XRatio = CDbl(SclX)
End Property

Public Property Get YRatio() As Double
    YRatio = CDbl(SclY)
End Property

Private Sub CopyVals()
    With cXAxis
        cXVals = .Data()
        cXMin = .Min
        cXMax = .Max
        cXAxisLarge = .TickLarge
        cXAxisSmall = .TickSmall
        cXAxisNumber = .TickNumber
    End With
    With cYAxis
        cYVals = .Data()
        cYMin = .Min
        cYMax = .Max
        cYAxisLarge = .TickLarge
        cYAxisSmall = .TickSmall
        cYAxisNumber = .TickNumber
    End With
End Sub

Private Sub DrawCanvas()
Static BrdX, BrdY0, BrdY1, k As Single, m As Single
Static MrgUp As Integer, MrgDown As Integer, MrgLeft As Integer, MrgRight As Integer
Static cFacX, cFacY, cCor As Single
Static BTikY, KTikY, BTikX, KTikX

Static oSize As Long

    cObject.Cls
    MrgUp = 20 * TwpY
    MrgDown = 30 * TwpY
    MrgLeft = 35 * TwpX
    MrgRight = 20 * TwpX
    
    SclX = (cXMax - cXMin) / cObject.Width
    SclY = (cYMax - cYMin) / cObject.Height
    
    With ScaleArea
        .Left = cXMin - SclX * MrgLeft
        .Right = cXMax + SclX * MrgRight
        .Top = cYMax + SclY * MrgUp
        .Bottom = cYMin - SclY * MrgDown
        cObject.Scale (.Left, .Top)-(.Right, .Bottom)
    End With
    
    cObject.Line (cXMin, cYMin)-(cXMax, cYMax), , B
    
    cFacX = SclX * TwpX
    cFacY = SclY * TwpY
        
    BTikY = cFacY * 12: KTikY = cFacY * 8
    BTikX = cFacX * 12: KTikX = cFacX * 8
    
    ' --------------------------------- X Axis drawing -----------------------------------------------
    cCor = IIf(Int(cXAxisNumber) = 0, 10, 1)
    If cXMin < 0 Then
        For k = 0 To cXMin Step -cXAxisSmall
            cObject.Line (k, cYMin)-(k, (cYMin - KTikY)), vbBlack
            cObject.Line (k, cYMax)-(k, (cYMax + KTikY)), vbBlack
        Next
        
        For k = 0 To cXMin Step -cXAxisLarge
            GetTextExtentPoint32 cObject.hdc, Round(k), Len(CStr(Round(k))) + 2, TextSz
            cObject.CurrentX = k - cFacX * (TextSz.cx / 2)
            cObject.CurrentY = cYMin - BTikY
            If (Abs(k * cCor) Mod cXAxisNumber * cCor) = 0 Then cObject.Print Round(k)
                
            cObject.Line (k, cYMin)-(k, (cYMin - BTikY)), vbBlack
            cObject.Line (k, cYMax)-(k, (cYMax + BTikY)), vbBlack
        Next
    End If
    
    For k = 0 To cXMax Step cXAxisSmall
        cObject.Line (k, cYMin)-(k, (cYMin - KTikY)), vbBlack
        cObject.Line (k, cYMax)-(k, (cYMax + KTikY)), vbBlack
    Next
    
    For k = 0 To cXMax Step cXAxisLarge
        GetTextExtentPoint32 cObject.hdc, Round(k), Len(CStr(Round(k))) + 2, TextSz
        cObject.CurrentX = k - cFacX * (TextSz.cx / 2)
        cObject.CurrentY = cYMin - BTikY
        If (k * cCor Mod cXAxisNumber * cCor) = 0 Then cObject.Print Round(k)
            
        cObject.Line (k, cYMin)-(k, (cYMin - BTikY)), vbBlack
        cObject.Line (k, cYMax)-(k, (cYMax + BTikY)), vbBlack
    Next
    

'' ------------------------------------------ Y Axis Drawing -------------------------------------------
    cCor = IIf(Int(cYAxisNumber) = 0, 10, 1)
    If cYMin < 0 Then
        For k = 0 To cYMin Step -cYAxisSmall
            cObject.Line (cXMin, k)-((cXMin - KTikX), k), vbBlack
            cObject.Line (cXMax, k)-((cXMax + KTikX), k), vbBlack
        Next
        
        For k = 0 To cYMin Step -cYAxisLarge
            GetTextExtentPoint32 cObject.hdc, k, Len(CStr(k)) + 2, TextSz
            cObject.CurrentX = cXMin - BTikX - cFacX * (TextSz.cx)
            cObject.CurrentY = k + cFacY * 8
            If (Abs(k * cCor) Mod cYAxisNumber * cCor) = 0 Then cObject.Print Round(k, 1)
                
            cObject.Line (cXMin, k)-((cXMin - BTikX), k), vbBlack
            cObject.Line (cXMax, k)-((cXMax + BTikX), k), vbBlack
        Next k
    End If
    
    For k = 0 To cYMax Step cYAxisSmall
        cObject.Line (cXMin, k)-((cXMin - KTikX), k), vbBlack
        cObject.Line (cXMax, k)-((cXMax + KTikX), k), vbBlack
    Next
    
    For k = 0 To cYMax Step cYAxisLarge
        GetTextExtentPoint32 cObject.hdc, k, Len(CStr(k)) + 2, TextSz
        cObject.CurrentX = cXMin - BTikX - cFacX * (TextSz.cx)
        cObject.CurrentY = k + cFacY * 8
        If (k * cCor Mod cYAxisNumber * cCor) = 0 Then cObject.Print Round(k, 1)
            
        cObject.Line (cXMin, k)-((cXMin - BTikX), k), vbBlack
        cObject.Line (cXMax, k)-((cXMax + BTikX), k), vbBlack
    Next k
    
    With cObject
        oSize = .FontSize
        .FontSize = 12
        .FontBold = True
        GetTextExtentPoint32 .hdc, cFileName, Len(cFileName) + 4, TextSz
        .CurrentX = cXMax - cFacX * (TextSz.cx)
        .CurrentY = cYMax - cFacY * (TextSz.cy)
        cObject.Print cFileName
        .FontBold = False
        .FontSize = oSize
    End With
    
End Sub

Public Sub DrawGraph()
    cObject.Picture = Nothing
    Call CopyVals
    Call DrawCanvas
    
    For k = 0 To UBound(cXVals)
        cObject.Circle (cXVals(k), cYVals(k)), SclX * 24, DotColor
    Next k
    cObject.Line (0, cYMin)-(0, cYMax), AxisColor
    cObject.Picture = cObject.Image
    cObject.Line (0, cYMin)-(0, cYMax), vbBlue
End Sub

Private Sub Class_Initialize()
    TwpX = Screen.TwipsPerPixelX
    TwpY = Screen.TwipsPerPixelY
    Set cXAxis = New clsAxis
    Set cYAxis = New clsAxis
End Sub
