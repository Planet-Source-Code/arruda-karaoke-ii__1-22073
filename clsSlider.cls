VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Dim UnidVert As Integer, UnidHorz As Integer
Dim VolBarY1 As Integer, VolBarY2 As Integer, VolBarX1 As Integer, VolBarX2 As Integer, VolBarHeight As Integer
Dim BalanceBarY1 As Integer, BalanceBarY2 As Integer, BalanceBarX1 As Integer, BalanceBarX2 As Integer, BalanceBarHeight As Integer

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim HorzButton As RECT
Dim VertButton As RECT

Private mvarShowFrame As Boolean
Private mvarControl As PictureBox
Private mvarVolume As Integer
Private mvarBalance As Integer
Private mvarEnabled As Boolean
Private Sub DrawState(ByVal b As Boolean)
    
    Exit Sub
    Dim CrColor As Long
    If b Then
        CrColor = &H0
    Else
        CrColor = &H2F2F2F
    End If
    i = VolBarY1
    For j = 1 To 10
        Control.Line ((UnidHorz * 3) - 2, i)-Step(2, 0), CrColor
        Control.Line (VolBarX2, i)-Step(2, 0), CrColor
        i = i + VolBarHeight / 9
    Next
    Control.Line (BalanceBarX1, BalanceBarY1 + 6)-Step(0, 2), CrColor
    Control.Line (BalanceBarX1 + Int((BalanceBarHeight / 2)), BalanceBarY1 + 6)-Step(0, 2), CrColor
    Control.Line (BalanceBarX2 - 1, BalanceBarY1 + 6)-Step(0, 2), CrColor
    DrawButtonVert
    DrawButtonHorz
    
End Sub

Public Property Let Enabled(ByVal vData As Boolean)
    
    mvarEnabled = vData
    Control.Enabled = mvarEnabled
    DrawState mvarEnabled

End Property


Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property





Public Property Let Balance(ByVal vData As Integer)

    mvarBalance = IIf(vData < 0, 0, vData)
    mvarBalance = IIf(vData > 127, 127, vData)
    X = (((BalanceBarHeight / 127) * mvarBalance)) + (BalanceBarX1 - Int(((HorzButton.Right - HorzButton.Left) / 2)))
    HorzButton.Left = X
    HorzButton.Right = X + Int(UnidHorz * 1.5)
    Control.Cls
    DrawButtonHorz
    DrawButtonVert


End Property


Public Property Get Balance() As Integer
    Balance = mvarBalance
End Property



Public Property Let Volume(ByVal vData As Integer)
    
    mvarVolume = IIf(vData < 0, 0, vData)
    mvarVolume = IIf(vData > 127, 127, vData)
    Y = VolBarHeight - (((VolBarHeight / 127) * mvarVolume)) + (VolBarY1 - Int(((VertButton.Bottom - VertButton.Top) / 2)))
    VertButton.Top = Y
    VertButton.Bottom = Y + UnidVert
    Control.Cls
    DrawButtonHorz
    DrawButtonVert
    
End Property


Public Property Get Volume() As Integer
    Volume = mvarVolume
End Property



Private Property Set Control(ByVal vData As PictureBox)
    Set mvarControl = vData
End Property


Private Property Get Control() As PictureBox
    Set Control = mvarControl
End Property



Sub DragButton(ByVal Button As Integer, X As Single, Y As Single)

    If Y >= VolBarY1 Then
        If Button = 1 Then
            If ((X < VertButton.Left) Or (X > VertButton.Right)) Then
                ReleaseCapture
                Exit Sub
            End If
            SetCapture Control.hWnd
            Y = Y - 5
            Y = IIf(Y + Int(((VertButton.Bottom - VertButton.Top) / 2)) > VolBarY2, VolBarY2 - Int(((VertButton.Bottom - VertButton.Top) / 2)), Y)
            Y = IIf(Y + Int(((VertButton.Bottom - VertButton.Top) / 2)) < VolBarY1, VolBarY1 - Int(((VertButton.Bottom - VertButton.Top) / 2)), Y)
            VertButton.Top = Y
            VertButton.Bottom = Y + UnidVert
            mvarVolume = 127 - ((127 / VolBarHeight) * ((Y - VolBarY1) + Int(((VertButton.Bottom - VertButton.Top) / 2))))
        Else
            ReleaseCapture
        End If
    Else
        If Button = 1 Then
            If ((Y < HorzButton.Top) Or (Y > HorzButton.Bottom)) Then
                ReleaseCapture
                Exit Sub
            End If
            SetCapture Control.hWnd
            X = X - 3
            X = IIf(X + Int(((HorzButton.Right - HorzButton.Left) / 2)) > BalanceBarX2, BalanceBarX2 - Int(((HorzButton.Right - HorzButton.Left) / 2)), X)
            X = IIf(X + Int(((HorzButton.Right - HorzButton.Left) / 2)) < BalanceBarX1, BalanceBarX1 - Int(((HorzButton.Right - HorzButton.Left) / 2)), X)
            HorzButton.Left = X
            HorzButton.Right = X + Int(UnidHorz * 1.5)
            mvarBalance = ((127 / BalanceBarHeight) * ((X - BalanceBarX1) + Int(((HorzButton.Right - HorzButton.Left) / 2))))
            Control.Cls
        Else
            ReleaseCapture
        End If
    End If
    Control.Cls
    DrawButtonVert
    DrawButtonHorz
    
End Sub

Private Sub DrawButtonVert()
    
    DrawFrameControl Control.hdc, VertButton, 4, &H10
    
End Sub

Private Sub DrawButtonHorz()

    DrawFrameControl Control.hdc, HorzButton, 4, &H10

End Sub

Public Property Let ShowFrame(ByVal vData As Boolean)
    
    mvarShowFrame = vData
    Control.Cls
    If ShowFrame Then
        Control.Line (0, 0)-Step(Control.ScaleWidth - 2, Control.ScaleHeight - 2), vb3DShadow, B
        Control.Line (1, 1)-Step(Control.ScaleWidth - 2, Control.ScaleHeight - 2), &HC0C0C0, B
    End If
    Control.Picture = Control.Image
    Volume = 0
    Balance = 63
    DrawButtonVert
    DrawButtonHorz

End Property


Public Property Get ShowFrame() As Boolean
    ShowFrame = mvarShowFrame
End Property




Sub SlideCreate(ByVal Pic As PictureBox)

    
    Set Control = Pic
    Control.ScaleMode = vbPixels
    Control.AutoRedraw = True
    Control.BorderStyle = 0
    
    
    UnidVert = Control.ScaleHeight / 10
    UnidHorz = Control.ScaleWidth / 10
    
    VolBarY1 = Control.ScaleHeight / 3
    VolBarY2 = Control.ScaleHeight - UnidVert
    VolBarX1 = (Control.ScaleWidth / 2) - 1
    VolBarX2 = Control.ScaleWidth - (UnidHorz * 3)
    VolBarHeight = VolBarY2 - VolBarY1
    Control.Line (VolBarX1, VolBarY1 - 1)-Step(0, VolBarHeight + 2), &H0
    Control.Line (VolBarX1 + 1, VolBarY1 - 1)-Step(0, VolBarHeight + 2), &HC0C0C0
    
    i = VolBarY1
    For j = 1 To 10
        Control.Line ((UnidHorz * 3) - 2, i)-Step(2, 0), &H0
        Control.Line ((UnidHorz * 3) - 2, i + 1)-Step(2, 0), &HC0C0C0
        Control.Line (VolBarX2, i)-Step(2, 0), &H0
        Control.Line (VolBarX2, i + 1)-Step(2, 0), &HC0C0C0
        i = i + VolBarHeight / 9
    Next
    
    BalanceBarY1 = ((VolBarY1) / 2) - UnidVert / 2
    BalanceBarY2 = ((VolBarY1) / 2) - UnidVert / 2
    BalanceBarX1 = UnidHorz * 2
    BalanceBarX2 = Control.ScaleWidth - UnidHorz * 2
    BalanceBarHeight = BalanceBarX2 - BalanceBarX1
    Control.Line (BalanceBarX1, BalanceBarY1 + 1)-Step(BalanceBarHeight, 0), &HC0C0C0
    Control.Line (BalanceBarX1, BalanceBarY1)-Step(BalanceBarHeight, 0), &H0
    
    Control.Line (BalanceBarX1, BalanceBarY1 + 6)-Step(0, 2), &H0
    Control.Line (BalanceBarX1 + 1, BalanceBarY1 + 6)-Step(0, 2), &HC0C0C0

    Control.Line (BalanceBarX1 + Int((BalanceBarHeight / 2)), BalanceBarY1 + 6)-Step(0, 2), &H0
    Control.Line (BalanceBarX1 + Int((BalanceBarHeight / 2)) + 1, BalanceBarY1 + 6)-Step(0, 2), &HC0C0C0

    Control.Line (BalanceBarX2 - 1, BalanceBarY1 + 6)-Step(0, 2), &H0
    Control.Line (BalanceBarX2, BalanceBarY1 + 6)-Step(0, 2), &HC0C0C0

    Control.Picture = Control.Image
    
    VertButton.Left = VolBarX1 - (UnidHorz * 2) + 1
    VertButton.Right = VertButton.Left + (UnidHorz * 4)
    VertButton.Bottom = VolBarY1 + UnidVert
    VertButton.Top = VolBarY1

    HorzButton.Right = BalanceBarX1 + Int(UnidHorz * 1.5)
    HorzButton.Left = BalanceBarX1
    HorzButton.Top = Int(UnidVert / 1.5)
    Bt = HorzButton.Top + Int(UnidVert * 2)
    If Bt > BalanceBarY1 + 10 Then Bt = BalanceBarY1 + 10
    HorzButton.Bottom = Bt
    ShowFrame = False

End Sub




