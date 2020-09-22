VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSliderVert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Private mvarValue As Integer
Private mvarBalance As Integer
Private mvarEnabled As Boolean
Private Sub DrawState(ByVal b As Boolean)
    
    Dim CrColor As Long
    If b Then
        CrColor = &H0
    Else
        CrColor = vb3DShadow
    End If
    i = VolBarY1
    i = VolBarY1
    For j = 1 To 10
        Control.Line ((UnidHorz * 3) - 2, i)-Step(2, 0), CrColor
        Control.Line (VolBarX2, i)-Step(2, 0), CrColor
        i = i + VolBarHeight / 9
    Next
    DrawButtonVert

End Sub

Public Property Let Enabled(ByVal vData As Boolean)
    
    mvarEnabled = vData
    Control.Enabled = mvarEnabled
    DrawState mvarEnabled

End Property


Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property





Public Property Let Value(ByVal vData As Integer)
    
    mvarValue = IIf(vData < 0, 0, vData)
    mvarValue = IIf(vData > 127, 127, vData)
    Y = VolBarHeight - (((VolBarHeight / 127) * mvarValue)) + (VolBarY1 - Int(((VertButton.Bottom - VertButton.Top) / 2)))
    VertButton.Top = Y
    VertButton.Bottom = Y + UnidVert
    Control.Cls
    DrawButtonVert
    
End Property


Public Property Get Value() As Integer
    Value = mvarValue
End Property



Private Property Set Control(ByVal vData As PictureBox)
    Set mvarControl = vData
End Property


Private Property Get Control() As PictureBox
    Set Control = mvarControl
End Property



Sub DragButton(ByVal Button As Integer, X As Single, Y As Single)

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
        mvarValue = 127 - ((127 / VolBarHeight) * ((Y - VolBarY1) + Int(((VertButton.Bottom - VertButton.Top) / 2))))
    Else
        ReleaseCapture
    End If
    Control.Cls
    DrawButtonVert
    
End Sub

Private Sub DrawButtonVert()
    
    DrawFrameControl Control.hdc, VertButton, 4, &H10
    
End Sub

Public Property Let ShowFrame(ByVal vData As Boolean)
    
    mvarShowFrame = vData
    Control.Cls
    If ShowFrame Then
        Control.Line (0, 0)-Step(Control.ScaleWidth - 2, Control.ScaleHeight - 2), vb3DShadow, B
        Control.Line (1, 1)-Step(Control.ScaleWidth - 2, Control.ScaleHeight - 2), vb3DHighlight, B
    End If
    Control.Picture = Control.Image
    Value = 0
    DrawButtonVert

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
    
    VolBarY1 = UnidVert
    VolBarY2 = Control.ScaleHeight - UnidVert
    VolBarX1 = (Control.ScaleWidth / 2) - 1
    VolBarX2 = Control.ScaleWidth - (UnidHorz * 3)
    VolBarHeight = VolBarY2 - VolBarY1
    Control.Line (VolBarX1, VolBarY1 - 1)-Step(0, VolBarHeight + 2), &H0
    Control.Line (VolBarX1 + 1, VolBarY1 - 1)-Step(0, VolBarHeight + 2), &HFFFFFF
    
    i = VolBarY1
    For j = 1 To 10
        Control.Line ((UnidHorz * 3) - 2, i)-Step(2, 0), &H0
        Control.Line ((UnidHorz * 3) - 2, i + 1)-Step(2, 0), vb3DHighlight
        Control.Line (VolBarX2, i)-Step(2, 0), &H0
        Control.Line (VolBarX2, i + 1)-Step(2, 0), vb3DHighlight
        i = i + VolBarHeight / 9
    Next

    VertButton.Left = VolBarX1 - (UnidHorz * 2) + 1
    VertButton.Right = VertButton.Left + (UnidHorz * 4) + 1
    VertButton.Bottom = VolBarY1 + UnidVert
    VertButton.Top = VolBarY1
    Control.Picture = Control.Image
    ShowFrame = False

End Sub

