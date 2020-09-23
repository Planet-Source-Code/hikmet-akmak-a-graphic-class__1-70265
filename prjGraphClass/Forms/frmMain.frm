VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGraphEx 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   5340
      Left            =   225
      ScaleHeight     =   5310
      ScaleWidth      =   7650
      TabIndex        =   0
      Top             =   225
      Width           =   7680
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MinX As Double, MinY As Double, MaxX As Double, MaxY As Double
Private CD As cDlg, SepChr(0 To 2) As String, AllData As String
Private cGraph As clsGraph
Private Fso As FileSystemObject

Private Sub Form_Load()
Static p As Long
Static SplitData() As String
Static xVals() As Double, yVals() As Double
Static NumDat As Long
Static SpCh As String

    Set CD = New cDlg
    CD.hOwner = Me.hWnd
    
    Set cGraph = New clsGraph
    Set Fso = New FileSystemObject
    
    SepChr(0) = Chr(32)      'space character
    SepChr(1) = Chr(9)      'tab character
    SepChr(2) = Chr(44)      'coma character
    
    'Read data first"
    Call ReadAllFile(App.Path & "\ExampleData.txt")
    
    'if no data read exit sub
    If Len(AllData) = 0 Then Exit Sub
    'Split data with vbCrLf
    SplitData = Split(AllData, vbCrLf)
    NumDat = UBound(SplitData)
    
    ReDim xVals(NumDat), yVals(NumDat)
    'Find data seperation character
    SpCh = SpacedChr(SplitData(0))
    
    p = 0
    Do
        If Len(SplitData(p)) > 0 Then
            xVals(p) = Val(Split(SplitData(p), SpCh)(0))
            yVals(p) = Val(Split(SplitData(p), SpCh)(1))
            MaxX = FindMax(xVals(p), MaxX)
            MinX = FindMin(xVals(p), MinX)
            MaxY = FindMax(yVals(p), MaxY)
            MinY = FindMin(yVals(p), MinY)
        End If
        p = p + 1
    Loop Until (p > NumDat)
            
    'transfer data to graph class
    With cGraph
        .GraphObject = picGraphEx
        .FileName = "ExampleData.txt"
        With .XAxis
            .Data = xVals()
            .Max = MaxX
            .Min = MinX
            .TickLarge = 5
            .TickSmall = 1
            .TickNumber = 10
        End With
        With .YAxis
            .Data = yVals()
            .Max = MaxY
            .Min = MinY
            .TickLarge = 0.2
            .TickSmall = 0.04
            .TickNumber = 0.2
        End With
        .DrawGraph              'draw graph
    End With

End Sub

Private Function SpacedChr(sVal As String) As String
Static m As Integer
    For m = 0 To 2
        If InStr(1, sVal, SepChr(m)) <> 0 Then
            SpacedChr = SepChr(m)
            Exit For
        End If
    Next m
End Function

Private Sub ReadAllFile(sFile As String)
Static aff As Long
    
    If Fso.FileExists(sFile) Then
        aff = FreeFile
        Open sFile For Binary As #aff
        AllData = Space(LOF(aff))
        Get #aff, , AllData
        Close #aff
    End If

End Sub

Private Function FindMax(X, Y)
    FindMax = X
    If Y > X Then FindMax = Y
End Function

Private Function FindMin(X, Y)
    FindMin = X
    If Y < X Then FindMin = Y
End Function

Private Sub Form_Resize()
    If Me.ScaleWidth > 0 And Me.ScaleHeight > 0 Then
        picGraphEx.Move 30, 30, Me.ScaleWidth - 60, Me.ScaleHeight - 60
        cGraph.DrawGraph
    End If
End Sub
