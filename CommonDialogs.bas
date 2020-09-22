Attribute VB_Name = "CommonDialogs"
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function ComDlgFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function VarPtr Lib "msvbvm50.dll" (var As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const LOGPIXELSY = 90
Const LF_FACESIZE = 32
Const CF_EFFECTS = &H100&
Const CF_TTONLY = &H40000
Const CF_FORCEFONTEXIST = &H10000
Const REGULAR_FONTTYPE = &H400
Const CF_SCREENFONTS = &H1
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_USESTYLE = &H80&
Const OFN_OVERWRITEPROMPT = &H2
Const OFN_HIDEREADONLY = &H4
Const OFN_CREATEPROMPT = &H2000
Const OFN_FILEMUSTEXIST = &H1000

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
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

Public Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long
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


Private Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type CHOOSEFONTRETURN
    lfBold As Long
    lfItalic As Byte
    lfSize As Integer
    lfFaceName As String
    lfOK As Long
End Type

Private Type LOGFONT
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

Public Type NEW_FONT
        FontName As String
        FontBold As Boolean
        FontItalic As Boolean
        FontSize As Single
End Type
Global NewFont As NEW_FONT
Public Function ShowFont(Frm As Form, FtName As String, Optional FtStyle As String, Optional FtSize As Single) As Long
    
    Dim cf As CHOOSEFONT
    Dim Lf As LOGFONT
    Dim Rf As CHOOSEFONTRETURN
    Dim i As Integer
    Dim retVal As Long
    
    cf.lStructSize = Len(cf)
    cf.hwndOwner = Frm.hWnd
    cf.lpLogFont = VarPtr(Lf)
    cf.rgbColors = 0
    cf.flags = CF_SCREENFONTS Or CF_INITTOLOGFONTSTRUCT Or CF_TTONLY

    For i = 0 To Len(FtName) - 1
        Lf.lfFaceName(i) = Asc(Mid(FtName, i + 1, 1))
    Next
    
    If FtSize > 0 Then
        Px = (FtSize * GetDeviceCaps(Frm.hdc, LOGPIXELSY)) / 72
        Lf.lfHeight = Px
    End If
    If Trim(FtStyle) <> "" Then
        cf.lpszStyle = FtStyle
    End If
    
    If ComDlgFont(cf) = 1 Then Rf.lfOK = 1 Else Rf.lfOK = 0
    
    Rf.lfSize = cf.iPointSize / 10
    Rf.lfItalic = Lf.lfItalic
    Rf.lfBold = IIf(Lf.lfWeight > 400, 1, 0)
    
    For i = 0 To 31
        Rf.lfFaceName = Rf.lfFaceName + Chr(Lf.lfFaceName(i))
    Next
    
    Rf.lfFaceName = Mid(Rf.lfFaceName, 1, InStr(1, Rf.lfFaceName, Chr(0)) - 1)
    If Rf.lfSize > 0 Then
        NewFont.FontName = Rf.lfFaceName
        NewFont.FontBold = Rf.lfBold
        NewFont.FontItalic = Rf.lfItalic
        NewFont.FontSize = Rf.lfSize
        ShowFont = 1
    Else
        ShowFont = 0
    End If
    
End Function
Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
     
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BrowseInfo

    udtBI.hwndOwner = hwndOwner
    udtBI.lpszTitle = lstrcat(sPrompt, "")
    udtBI.ulFlags = BIF_RETURNONLYFSDIRS

     lpIDList = SHBrowseForFolder(udtBI)
     If lpIDList Then
          sPath = String$(260, 0)
          lResult = SHGetPathFromIDList(lpIDList, sPath)
          Call CoTaskMemFree(lpIDList)
          iNull = InStr(sPath, vbNullChar)
          If iNull Then
               sPath = Left$(sPath, iNull - 1)
          End If
     End If
     BrowseForFolder = sPath

End Function
Public Function ShowOpen(ByVal hForm As Long, Filter As String, Title As String, InitDir As String) As String
 
 Dim ofn As OPENFILENAME
    Dim a As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hForm
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = Title
    ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    a = GetOpenFileName(ofn)

    If (a) Then
        ShowOpen = Trim$(ofn.lpstrFile)
    Else
        ShowOpen = ""
    End If

End Function

Public Function ShowSave(ByVal hForm As Long, FileName As String, Filter As String, Title As String, InitDir As String) As String
 
    Dim ofn As OPENFILENAME
    Dim a As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hForm
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next

    ofn.lpstrFilter = Filter
    ofn.lpstrFile = FileName & Space$(255 - Len(FileName))
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = Title
    ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
    a = GetSaveFileName(ofn)

    If (a) Then
        ShowSave = Trim$(ofn.lpstrFile)
    Else
        ShowSave = ""
    End If

End Function


Function ShowColor(hForm As Long) As Long

    Dim Cc As ChooseColor
    Dim CustColor(16) As Long

    Cc.lStructSize = Len(Cc)
    Cc.hwndOwner = hForm
    Cc.hInstance = App.hInstance
    Cc.flags = 0
    Cc.lpCustColors = String$(16 * 4, 0)
    Dim a
    Dim X
    Dim c1
    Dim c2
    Dim c3
    Dim c4
    a = ChooseColor(Cc)
    
    If (a) Then
        ShowColor = Str$(Cc.rgbResult)
        For X = 1 To Len(Cc.lpCustColors) Step 4
            c1 = Asc(Mid$(Cc.lpCustColors, X, 1))
            c2 = Asc(Mid$(Cc.lpCustColors, X + 1, 1))
            c3 = Asc(Mid$(Cc.lpCustColors, X + 2, 1))
            c4 = Asc(Mid$(Cc.lpCustColors, X + 3, 1))
            CustColor(X / 4) = (c1) + (c2 * 256) + (c3 * 65536) + (c4 * 16777216)
        Next X
    Else
        ShowColor = -1
    End If

End Function


