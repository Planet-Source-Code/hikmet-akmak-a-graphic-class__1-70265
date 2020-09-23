VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAxis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cData() As Double
Private cMin As Double
Private cMax As Double
Private cTickLarge As Single
Private cTickSmall As Single
Private cTickNumber As Single
Private cLabel As String

Public Property Get Data() As Double()
    Data = cData
End Property

Public Property Let Data(ByRef mVal() As Double)
    cData = mVal
End Property

Public Property Get Min() As Double
    Min = cMin
End Property

Public Property Let Min(ByVal mVal As Double)
    cMin = NearestInt(mVal, 1)
End Property

Public Property Get Max() As Double
    Max = cMax
End Property

Public Property Let Max(ByVal mVal As Double)
    cMax = NearestInt(mVal, 1)
End Property

Public Property Get TickLarge() As Double
    TickLarge = cTickLarge
End Property

Public Property Let TickLarge(ByVal mVal As Double)
    cTickLarge = mVal
End Property

Public Property Get TickSmall() As Double
    TickSmall = cTickSmall
End Property

Public Property Let TickSmall(ByVal mVal As Double)
    cTickSmall = mVal
End Property

Public Property Get TickNumber() As Double
    TickNumber = cTickNumber
End Property

Public Property Let TickNumber(ByVal mVal As Double)
    cTickNumber = mVal
End Property

Public Property Get Label() As String
    Label = cLabel
End Property

Public Property Let Label(ByVal mVal As String)
    cLabel = mVal
End Property

Private Function NearestInt(sVal As Double, cDigit As Integer) As Double
Static vSgn As Integer, nDgt As Integer, fVal As Double

    nDgt = 10 ^ cDigit
    vSgn = Sgn(sVal)
    fVal = Abs(sVal) * nDgt
    
    Select Case (fVal - Int(fVal))
        Case Is = 0: NearestInt = vSgn * (fVal / nDgt)
        Case Is > 0: NearestInt = vSgn * ((Int(fVal) + 1) / nDgt)
    End Select
    
End Function

