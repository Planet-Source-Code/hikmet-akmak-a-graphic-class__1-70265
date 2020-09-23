VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'<*#*>This class was modified by russell sanders to show the open file, save file and browse for folder dialogs
'<*CATEGORIES*>

'<*PROPERTIES*>
'<*START*>
'<*PUBLIC_DEC*>
'<*FORM_DEC*>
'<*CODE*>

' this is a striped down version of the code at
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=9260&lngWId=1
'
Private Type OPENFILENAME 'Open & Save Dialog
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type CHOOSECOLOR 'Color Dialog
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    RGBResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Const LF_FACESIZE = 32 'Font Dialog
Private Type LOGFONT 'Font Dialog
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type ChooseFont 'Font Dialog
    lStructSize As Long
    hWndOwner As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type
' extra font constant
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const SCREEN_FONTTYPE = &H2000
Const BOLD_FONTTYPE = &H100
Const FW_BOLD = 700
Const LOGPIXELSY = 90

'------- Dialog calling functions
' -------------- Standard
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hOwner As Long) As Long
Private Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" (p1 As Any, p2 As Any) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Enum CdlgExt_Actions
       cdlgOpen = 1
       cdlgSave = 2
       cdlgColor = 3
       cdlgFont = 4
 End Enum
 
Public Enum CdlgExt_Flags
 ' Open & Save Dialog
 cdlOFNAllowMultiselect = &H200
 cdlOFNCreatePrompt = &H2000
 cdlOFNExplorer = &H80000
 cdlOFNExtensionDifferent = &H400
 cdlOFNFileMustExist = &H1000
 cdlOFNHelpButton = &H10
 cdlOFNHideReadOnly = &H4
 cdlOFNLongNames = &H200000
 cdlOFNNoChangeDir = &H8
 cdlOFNNoDereferenceLinks = &H100000
 cdlOFNNoLongNames = &H40000
 cdlOFNNoReadOnlyReturn = &H8000
 cdlOFNNoValidate = &H100
 cdlOFNOverwritePrompt = &H2
 cdlOFNPathMustExist = &H800
 cdlOFNReadOnly = &H1
 cdlOFNShareAware = &H4000
 'Color Dialog
 cdlCCANYCOLOR = &H100
 cdlCCFullOpen = &H2
 cdlCCHelpButton = &H8
 cdlCCPreventFullOpen = &H4
 cdlCCRGBInit = &H1
' Font Dialog
 cdlCFANSIOnly = &H400
 cdlCFApply = &H200
 cdlCFBoth = &H3
 cdlCFEffects = &H100
 cdlCFFixedPitchOnly = &H4000
 cdlCFForceFontExist = &H10000
 cdlCFHelpButton = &H4
 cdlCFLimitSize = &H2000
 cdlCFNoFaceSel = &H80000
 cdlCFNoSimulations = &H1000
 cdlCFNoSizeSel = &H200000
 cdlCFNoStyleSel = &H100000
 cdlCFNoVectorFonts = &H800
 cdlCFPrinterFonts = &H2
 cdlCFScalableOnly = &H20000
 cdlCFScreenFonts = &H1
 cdlCFTTOnly = &H40000
 cdlCFWYSIWYG = &H8000
End Enum

Private RetValue As Long 'General
Const MAX_PATH = 260 'General
Private OFN As OPENFILENAME ' Open & Save Dialog

'Inner variables for properties
Private mFileName As String
Private mFileTitle As String
Private mhOwner As Long
Private mDialogTitle As String
Private mFilter As String
Private mInitDir As String
Private mSelDir As String
Private mDefaultExt As String
Private mFilterIndex As Long
Private mHelpFile As String
Private mHelpKey As Long
Private mRGBResult As Long
Private mItalic As Boolean
Private mUnderline As Boolean
Private mStrikethru As Boolean
Private mFontName As String
Private mFontSize As Long
Private mBold As Boolean
Private mDialogPrompt As String
Private mFlags As CdlgExt_Flags
Private mCancelError As Boolean
Private mhIcon As Long
Private mAppName As String
Private i As Integer

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Function BrowseForFolder(Optional sCaption As String = "Select a folder", Optional sDefault As String = "") As String
    Const BIF_RETURNONLYFSDIRS = 1
    Const MAX_PATH = 260
    Dim lPos As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, tBrowse As BrowseInfo

    With tBrowse
        'Set the owner window
        .hWndOwner = mhOwner        'Me.hWnd in VB
        .lpszTitle = sCaption
        .ulFlags = BIF_RETURNONLYFSDIRS     'Return only if the user selected a directory
    End With

    'Show the dialog
    lpIDList = SHBrowseForFolder(tBrowse)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        lPos = InStr(sPath, vbNullChar)
        If lPos Then
            BrowseForFolder = Left$(sPath, lPos - 1)
            If Right$(BrowseForFolder, 1) <> "\" Then
                BrowseForFolder = BrowseForFolder & "\"
            End If
        End If
    Else
        'User cancelled, return default path
        BrowseForFolder = sDefault
    End If
End Function

Public Property Let CancelError(ByVal vData As Boolean)
   mCancelError = vData
End Property

Public Property Get CancelError() As Boolean
  CancelError = mCancelError
End Property

Public Property Get hOwner() As Long
    hOwner = mhOwner
End Property

Public Property Let hOwner(ByVal New_hOwner As Long)
    mhOwner = New_hOwner
End Property

Public Property Get flags() As CdlgExt_Flags
    flags = mFlags
End Property

Public Property Let flags(ByVal New_Flags As CdlgExt_Flags)
    mFlags = New_Flags
End Property

Public Property Get DialogTitle() As String
   DialogTitle = mDialogTitle
End Property

Public Property Let DialogTitle(sTitle As String)
   mDialogTitle = sTitle
End Property

Public Property Get DialogPrompt() As String
    DialogPrompt = mDialogPrompt
End Property

Public Property Let DialogPrompt(ByVal New_Prompt As String)
    mDialogPrompt = New_Prompt
End Property

Public Property Get AppName() As String
    AppName = mAppName
End Property

Public Property Let AppName(ByVal New_AppName As String)
    mAppName = New_AppName
End Property

Public Property Let hIcon(ByVal vData As Long)
    mhIcon = vData
End Property

Public Property Get hIcon() As Long
   hIcon = mhIcon
End Property

' Font Properties
Public Property Get FontBold() As Boolean
  FontBold = mBold
End Property

Public Property Let FontBold(bBold As Boolean)
   mBold = bBold
End Property

Public Property Get FontName() As String
   FontName = mFontName
End Property

Public Property Let FontName(sName As String)
   mFontName = sName
End Property

Public Property Get FontSize() As Long
  FontSize = mFontSize
End Property

Public Property Let FontSize(lSize As Long)
   mFontSize = lSize
End Property

Public Property Get FontItalic() As Boolean
  FontItalic = mItalic
End Property

Public Property Let FontItalic(BItalic As Boolean)
   mItalic = BItalic
End Property

Public Property Get FontStrikeThru() As Boolean
   FontStrikeThru = mStrikethru
End Property

Public Property Let FontStrikeThru(bStrikethru As Boolean)
   mStrikethru = bStrikethru
End Property

Public Property Get FontUnderline() As Boolean
   FontUnderline = mUnderline
End Property

Public Property Let FontUnderline(bUnderline As Boolean)
   mUnderline = bUnderline
End Property

' Open , Save
Public Property Get DefaultExt() As String
   DefaultExt = mDefaultExt
End Property

Public Property Let DefaultExt(sDefExt As String)
   mDefaultExt = sDefExt
End Property

Public Property Get FileName() As String
   FileName = mFileName
End Property

Public Property Let FileName(sFileName As String)
   mFileName = sFileName
End Property

Public Property Get FileTitle() As String
   FileTitle = mFileTitle
End Property

Public Property Let FileTitle(sTitle As String)
   mFileTitle = sTitle
End Property

Public Property Get Filter() As String
   Filter = mFilter
End Property

Public Property Let Filter(sFilter As String)
   mFilter = sFilter
End Property

Public Property Get FilterIndex() As Long
   FilterIndex = mFilterIndex
End Property

Public Property Let FilterIndex(lIndex As Long)
    mFilterIndex = lIndex
End Property

Public Property Get InitDir() As String
   InitDir = mInitDir
End Property

Public Property Let InitDir(sDir As String)
    mInitDir = sDir
End Property

Public Property Get SelDir() As String
   SelDir = mSelDir
End Property

Public Property Let SelDir(sDir As String)
    mSelDir = sDir
End Property

'Color Dialog
Public Property Get Color() As Long
   Color = mRGBResult
End Property

Public Property Let Color(lValue As Long)
   mRGBResult = lValue
End Property

Public Sub ShowOpen()
Dim iDelim As Integer
  InitOFN
  RetValue = GetOpenFileName(OFN)
    If RetValue > 0 Then
       iDelim = InStr(OFN.lpstrFileTitle, vbNullChar)
       If iDelim Then mFileTitle = Left$(OFN.lpstrFileTitle, iDelim - 1)
       iDelim = InStr(OFN.lpstrFile, vbNullChar)
       If iDelim Then mFileName = Left$(OFN.lpstrFile, iDelim - 1)
    Else
       If mCancelError Then Err.Raise 0
    End If
End Sub
Public Sub ShowSave()
Dim iDelim As Integer
  InitOFN
  RetValue = GetSaveFileName(OFN)
    If RetValue > 0 Then
       iDelim = InStr(OFN.lpstrFileTitle, vbNullChar)
       If iDelim Then mFileTitle = Left$(OFN.lpstrFileTitle, iDelim - 1)
       iDelim = InStr(OFN.lpstrFile, vbNullChar)
       If iDelim Then mFileName = Left$(OFN.lpstrFile, iDelim - 1)
    Else
       If mCancelError Then Err.Raise 0
    End If
End Sub
Private Sub InitOFN()
  Dim sTemp As String, i As Integer
  Dim uFlag As Long
  uFlag = mFlags And (&H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
  With OFN
       .lStructSize = Len(OFN)
       .hWndOwner = mhOwner
       .flags = uFlag
       .lpstrDefExt = mDefaultExt
       sTemp = mInitDir
       If sTemp = "" Then sTemp = App.Path
       .lpstrInitialDir = sTemp
       sTemp = mFileName
       .lpstrFile = sTemp & String$(255 - Len(sTemp), 0)
       .nMaxFile = 255
       .lpstrFileTitle = String$(255, 0)
       .nMaxFileTitle = 255
        sTemp = mFilter
            For i = 1 To Len(sTemp)
                If Mid(sTemp, i, 1) = "|" Then
                   Mid(sTemp, i, 1) = vbNullChar
                End If
            Next
        sTemp = sTemp & String$(2, 0)
        .lpstrFilter = sTemp
        .nFilterIndex = mFilterIndex
        .lpstrTitle = mDialogTitle
        .hInstance = App.hInstance
 End With
End Sub

Public Sub ShowColor()
Dim cc As CHOOSECOLOR
Dim a As Long
    cc.lStructSize = Len(cc)
    cc.hWndOwner = mhOwner
    cc.hInstance = App.hInstance
    cc.flags = mFlags And (&H1 Or &H2 Or &H4 Or &H8 Or &H100)
    cc.lpCustColors = GetSetting(App.EXEName, "Settings", "CustomColors", String$(16 * 4, 0)) 'load custom colors
    a = ChooseColorAPI(cc)
        If (a) Then
            mRGBResult = cc.RGBResult
            SaveSetting App.EXEName, "Settings", "CustomColors", cc.lpCustColors 'save custom colors
        Else
            If mCancelError Then Err.Raise (RetValue)
        End If
End Sub

Public Sub ShowFont()
  Dim CF As ChooseFont
  Dim LF As LOGFONT
  Dim TempByteArray() As Byte
  Dim ByteArrayLimit As Long
  Dim OldhDC As Long
  Dim FontToUse As Long
  Dim tbuf As String * 80
  Dim X As Long
  Dim uFlag As Long
  uFlag = mFlags And (&H1 Or &H2 Or &H3 Or &H4 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H10000 Or &H20000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
  TempByteArray = StrConv(mFontName & vbNullChar, vbFromUnicode)
  ByteArrayLimit = UBound(TempByteArray)
  With LF
        For X = 0 To ByteArrayLimit
           .lfFaceName(X) = TempByteArray(X)
        Next
    .lfHeight = mFontSize / 72 * GetDeviceCaps(GetDC(mhOwner), LOGPIXELSY)
    .lfItalic = mItalic * -1
    .lfUnderline = mUnderline * -1
    .lfStrikeOut = mStrikethru * -1
    If mBold Then .lfWeight = FW_BOLD
  End With
  With CF
      .lStructSize = Len(CF)
      .hWndOwner = mhOwner
      .hdc = GetDC(mhOwner)
      .lpLogFont = lstrcpy(LF, LF)
      If Not uFlag Then
         .flags = cdlCFScreenFonts
      Else
         .flags = uFlag Or cdlCFWYSIWYG
      End If
     .flags = .flags Or cdlCFEffects Or CF_INITTOLOGFONTSTRUCT
     .rgbColors = mRGBResult
     .lCustData = 0
     .lpfnHook = 0
     .lpTemplateName = 0
     .hInstance = 0
     .lpszStyle = 0
     .nFontType = SCREEN_FONTTYPE
     .nSizeMin = 0
     .nSizeMax = 0
     .iPointSize = mFontSize * 10
    End With
    RetValue = ChooseFont(CF)
    If RetValue = 0 Then
       If mCancelError Then Err.Raise (RetValue)
    Else
       With LF
            mItalic = .lfItalic * -1
            mUnderline = .lfUnderline * -1
            mStrikethru = .lfStrikeOut * -1
       End With
       With CF
            mFontSize = .iPointSize \ 10
            mRGBResult = .rgbColors
            If .nFontType And BOLD_FONTTYPE Then
                mBold = True
            Else
                mBold = False
            End If
       End With
       FontToUse = CreateFontIndirect(LF)
       If FontToUse = 0 Then Exit Sub
          OldhDC = SelectObject(CF.hdc, FontToUse)
          RetValue = GetTextFace(CF.hdc, 79, tbuf)
          mFontName = Mid$(tbuf, 1, RetValue)
       End If
End Sub


'<*END*>


