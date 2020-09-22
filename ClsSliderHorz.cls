VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSliderHorz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim HorzButton As RECT
Dim UnidVert As Integer, UnidHorz As Integer
Dim ValueBarY1 As Integer, ValueBarY2 As Integer, ValueBarX1 As Integer, ValueBarX2 As Integer, ValueBarWidth As Integer

Private mvarShowFrame As Boolean
Private mvarControl As PictureBox
Private mvarValue As Integer
Private mvarEnabled As Boolean
Private Sub DrawState(ByVal b As Boolean)
    
    Dim CrColor As Long
    If b Then
        CrColor = &H0
    Else
        CrColor = vb3DShadow
    End If
    i = ValueBarX1
    For j = 1 To 10
        Control.Line (i, ValueBarY2)-Step(0, 2), CrColor
        i = i + ValueBarWidth / 9
    Next
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





Public Property Let Value(ByVal vData As Integer)

    mvarValue = IIf(vData < 0, 0, vData)
    mvarValue = IIf(vData > 127, 127, vData)
    X = (((ValueBarWidth / 127) * mvarValue)) + (ValueBarX1 - Int(((HorzButton.Right - HorzButton.Left) / 2)))
    HorzButton.Left = X
    HorzButton.Right = X + UnidHorz
    Control.Cls
    DrawButtonHorz

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
        If ((Y < HorzButton.Top) Or (Y > HorzButton.Bottom)) Then
            ReleaseCapture
            Exit Sub
        End If
        SetCapture Control.hWnd
        X = X - 3
        X = IIf(X + Int(((HorzButton.Right - HorzButton.Left) / 2)) > ValueBarX2, ValueBarX2 - Int(((HorzButton.Right - HorzButton.Left) / 2)), X)
        X = IIf(X + Int(((HorzButton.Right - HorzButton.Left) / 2)) < ValueBarX1, ValueBarX1 - Int(((HorzButton.Right - HorzButton.Left) / 2)), X)
        HorzButton.Left = X
        HorzButton.Right = X + Int(UnidHorz)
        mvarValue = ((127 / ValueBarWidth) * ((X - ValueBarX1) + Int(((HorzButton.Right - HorzButton.Left) / 2))))
        Control.Cls
    Else
        ReleaseCapture
    End If
    Control.Cls
    DrawButtonHorz

End Sub

Private Sub DrawButtonHorz()

    DrawFrameControl Control.hdc, HorzButton, 4, &H10

End Sub

Public Property Let ShowFrame(ByVal vData As Boolean)
    
    mvarShowFrame = vData
    Control.Cls
    If ShowFrame Then
        Control.Line (0, 0)-Step(Control.ScaleWidth - 2, Control.ScaleHeight - 2), vb3DShadow, B
        Control.Line (1, 1)-Step(Control.ScaleWidth - 2, Control.ScaleHeight - 2), vb3DHighlight, B
    End If
    Control.Picture = Control.Image
    DrawButtonHorz
    Value = 0

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
    ValueBarY1 = (Control.ScaleHeight - (UnidVert * 1.5)) / 2
    ValueBarY2 = Control.ScaleHeight - (UnidVert * 3)
    ValueBarX1 = UnidHorz
    ValueBarX2 = Control.ScaleWidth - UnidHorz
    ValueBarWidth = ValueBarX2 - ValueBarX1
    Control.Line (ValueBarX1, ValueBarY1 + 1)-Step(ValueBarWidth, 0), &HFFFFFF
    Control.Line (ValueBarX1, ValueBarY1)-Step(ValueBarWidth, 0), &H0
    
    i = ValueBarX1
    For j = 1 To 10
        Control.Line (i, ValueBarY2)-Step(0, 2), &H0
        Control.Line (i + 1, ValueBarY2)-Step(0, 2), vb3DHighlight
        i = i + ValueBarWidth / 9
    Next
    Control.Picture = Control.Image
    HorzButton.Left = ValueBarX1
    HorzButton.Right = ValueBarX1 + (UnidHorz * 2)
    HorzButton.Top = ValueBarY1 - (UnidVert * 2.5)
    HorzButton.Bottom = ValueBarY1 + (UnidVert * 2.5)
    ShowFrame = False

End Sub

