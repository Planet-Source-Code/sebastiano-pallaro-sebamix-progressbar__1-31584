VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "smProgressBar"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' API DECLARATIONS ==================================================
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
'====================================================================

Private mvarPercent As Integer
Private mvarCaption As String
Private Picture As PictureBox
Private mvarProgressColor As OLE_COLOR
Private mvarForeColor As OLE_COLOR
Private mvarBackColor As OLE_COLOR
Private mvarMax As Long
Private mvarMin As Long
Private mvarValue As Long
Private mvarUsePercent As Boolean
Private mvarShowCaption As Boolean
Private mvarOrientation As SMOrientation
Private mvarPictureProgress As SMDrawStyles
Private mvarOldDrawMode As Byte
Private mvarTextBefore As String
Private mvarTextAfter As String
Private mvarTextMiddle As String


Public Enum SMOrientation
    SMHorizontal = 0
    SMVertical = 1
End Enum

Public Enum SMDrawStyles
    SM_Color_Gray_Progress = -1
    SM_MergePen = vbMergePen
    SM_MergePenNot = vbMergePenNot
    SM_MergeNotPen = vbMergeNotPen
    SM_NotXorPen = vbNotXorPen
    SM_MaskPen = vbMaskPen
    SM_NotMaskPen = vbNotMaskPen
    SM_XorPen = vbXorPen
    SM_Invert = vbInvert
    SM_MaskPenNot = vbMaskPenNot
    SM_MaskNotPen = vbMaskNotPen
    SM_NotMergePen = vbNotMergePen
End Enum

Public Event Progress(Value As Long, Percent As Integer)

Public Property Let TextMiddle(ByVal vData As String)
    mvarTextMiddle = vData
End Property
Public Property Get TextMiddle() As String
    TextMiddle = mvarTextMiddle
End Property

Public Property Let TextAfter(ByVal vData As String)
    mvarTextAfter = vData
End Property
Public Property Get TextAfter() As String
    TextAfter = mvarTextAfter
End Property

Public Property Let TextBefore(ByVal vData As String)
    mvarTextBefore = vData
End Property
Public Property Get TextBefore() As String
    TextBefore = mvarTextBefore
End Property

Public Property Get OldDrawMode() As Byte
    OldDrawMode = mvarOldDrawMode
End Property

Private Sub subConvertPictureToGrayScale()
    Dim lngX As Long
    Dim lngY As Long
    Dim lngColor As Long
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    Dim lngGrayScaleColor As Long
        
    For lngX = 0 To Picture.ScaleWidth
        For lngY = 0 To Picture.ScaleHeight
            ' Leggo il colore del pixel...
            lngColor = GetPixel(Picture.hdc, lngX, lngY)
            
            ' Lo converto in R.G.B.
            lngRed = lngColor And &HFF
            lngGreen = (lngColor \ &H100) And &HFF
            lngBlue = (lngColor \ &H10000) And &HFF
            
            ' Con i valori R.G.B. lo converto in grigio...
            'lngGrayScaleColor = ((lngRed * 0.222) + (lngGreen * 0.707) + (lngBlue + 0.071))
            lngGrayScaleColor = lngRed * 0.299 + lngGreen * 0.587 + lngBlue + 0.114
            
            ' Disegno il pixel grigio...
            SetPixel Picture.hdc, lngX, lngY, lngGrayScaleColor
        Next lngY
    Next lngX
End Sub

Public Property Let PictureProgress(ByVal vData As SMDrawStyles)
    Select Case vData
        Case SMDrawStyles.SM_Color_Gray_Progress
            subConvertPictureToGrayScale
        Case Else
            Picture.DrawMode = vData
    End Select
    mvarPictureProgress = vData
End Property
Public Property Get PictureProgress() As SMDrawStyles
    PictureProgress = mvarPictureProgress
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    mvarBackColor = vData
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Let Orientation(ByVal vData As SMOrientation)
    mvarOrientation = vData
End Property
Public Property Get Orientation() As SMOrientation
    Orientation = mvarOrientation
End Property

Public Property Let ShowCaption(ByVal vData As Boolean)
    mvarShowCaption = vData
End Property
Public Property Get ShowCaption() As Boolean
    ShowCaption = mvarShowCaption
End Property

Public Property Let UsePercent(ByVal vData As Boolean)
    mvarUsePercent = vData
End Property
Public Property Get UsePercent() As Boolean
    UsePercent = mvarUsePercent
End Property

Private Sub CalculatePercent()
    mvarPercent = (100 * mvarValue) / mvarMax
End Sub

Public Property Let Value(ByVal vData As Long)
    If vData < mvarMin Then Exit Property
    If vData > mvarMax Then Exit Property
    mvarValue = vData
    Picture.BackColor = mvarBackColor
    CalculatePercent
    If mvarUsePercent = True Then
        ProgressBarPercent
    Else
        ProgressBarValue
    End If
    RaiseEvent Progress(mvarValue, mvarPercent)
End Property
Public Property Get Value() As Long
    Value = mvarValue
End Property

Public Property Let Min(ByVal vData As Long)
    If vData < 0 Then Exit Property
    If vData >= mvarMax Then Exit Property
    mvarMin = vData
End Property
Public Property Get Min() As Long
    Min = mvarMin
End Property

Public Property Let Max(ByVal vData As Long)
    If vData < 0 Then Exit Property
    If vData <= mvarMin Then Exit Property
    mvarMax = vData
End Property
Public Property Get Max() As Long
    Max = mvarMax
End Property

Public Sub InitPB(ByVal myPictureBox As Object, Optional myOrientation As SMOrientation)
    Set Picture = myPictureBox
    mvarOrientation = myOrientation
    mvarBackColor = RGB(150, 150, 150)
    mvarProgressColor = RGB(100, 100, 250)
    mvarForeColor = vbWhite
    Picture.AutoRedraw = True
    'Picture.ScaleMode = vbTwips
    Picture.ScaleMode = vbPixels
    
    mvarOldDrawMode = Picture.DrawMode
    
    Value = mvarValue
End Sub

Private Sub ProgressBarPercent()
    Dim PicWidth As Single
    Dim PicHeight As Single
    Dim BarWidth As Single
    Dim BarHeight As Single
    Dim Txt As String
    Picture.Cls
    PicWidth = Picture.ScaleWidth
    PicHeight = Picture.ScaleHeight
    
    If mvarValue > mvarMin Then
        If mvarOrientation = SMHorizontal Then
            BarWidth = (PicWidth * mvarPercent) / 100
            Picture.Line (0, 0)-(BarWidth, PicHeight), mvarProgressColor, BF
            Picture.Line (BarWidth + 1, 0)-(BarWidth, PicHeight), mvarProgressColor, BF
        Else
        
        
            BarHeight = (PicHeight * mvarPercent) / 100
            BarHeight = PicHeight - BarHeight
            
            Picture.Line (0, PicHeight)-(PicWidth, BarHeight), mvarProgressColor, BF
        End If
    End If
    If mvarShowCaption = True Then
        If mvarCaption = "" Then
            Txt = mvarTextBefore & " " & CStr(mvarPercent) & "%" & " " & mvarTextAfter
            Picture.CurrentX = (PicWidth - Picture.TextWidth(Txt)) / 2
            Picture.CurrentY = (PicHeight - Picture.TextHeight(Txt)) / 2
        Else
            Txt = mvarCaption
            Picture.CurrentX = (PicWidth - Picture.TextWidth(Txt)) / 2
            Picture.CurrentY = (PicHeight - Picture.TextHeight(Txt)) / 2
        End If
        Picture.ForeColor = mvarForeColor
        Picture.Print Txt
    End If
End Sub

Private Sub ProgressBarValue()
    Dim PicWidth As Single
    Dim PicHeight As Single
    Dim BarWidth As Single
    Dim BarHeight As Single
    Dim Txt As String
    
    Dim Val As Single
    
    Picture.Cls
    
    PicWidth = Picture.ScaleWidth
    PicHeight = Picture.ScaleHeight
    Val = mvarValue - mvarMin
    
    If mvarValue > mvarMin Then
        If mvarOrientation = SMHorizontal Then
            BarWidth = (PicWidth * Val) / (mvarMax - mvarMin)
            Picture.Line (0, 0)-(BarWidth, PicHeight), mvarProgressColor, BF
            Picture.Line (BarWidth + 1, 0)-(BarWidth, PicHeight), mvarProgressColor, BF
        Else
        
            'BarHeight = (PicHeight * mvarPercent) / 100
            BarHeight = (PicHeight * Val) / (mvarMax - mvarMin)
            BarHeight = PicHeight - BarHeight
        
            Picture.Line (0, PicHeight)-(PicWidth, BarHeight), mvarProgressColor, BF
            'Picture.Line (0, 0)-(PicWidth, BarHeight), mvarProgressColor, BF
            'Picture.Line (0, BarHeight + 1)-(PicWidth, BarHeight), mvarProgressColor, BF
        End If
    End If
    If mvarShowCaption = True Then
        If mvarCaption = "" Then
            Txt = mvarTextBefore & " " & CStr(mvarValue) & mvarTextMiddle & CStr(mvarMax) & " " & mvarTextAfter
        Else
            Txt = mvarCaption
        End If
        'If mvarOrientation = SMHorizontal Then
            Picture.CurrentX = (PicWidth - Picture.TextWidth(Txt)) / 2
            Picture.CurrentY = (PicHeight - Picture.TextHeight(Txt)) / 2
            Picture.ForeColor = mvarForeColor
            'Dim lngOldDM As Long
            'lngOldDM = Picture.DrawMode
            'Picture.DrawMode = 10
            Picture.Print Txt
            'Picture.DrawMode = lngOldDM
        'Else
        '    Dim x As Single
        '    Dim y As Single
        '    x = (PicWidth - Picture.TextWidth(Txt)) / 2
        '    y = (PicHeight - Picture.TextHeight(Txt)) / 2
        '    subTestoVerticale Picture, Txt, x, y
        'End If
    End If
End Sub

Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property
Public Property Get Caption() As String
    Caption = mvarCaption
End Property

Public Property Get Percent() As Integer
    Percent = mvarPercent
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    mvarForeColor = vData
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mvarForeColor
End Property

Public Property Let ProgressColor(ByVal vData As OLE_COLOR)
    mvarProgressColor = vData
End Property
Public Property Get ProgressColor() As OLE_COLOR
    ProgressColor = mvarProgressColor
End Property

Private Sub Class_Initialize()
    mvarCaption = ""
    mvarTextBefore = ""
    mvarTextAfter = ""
    mvarTextMiddle = "/"
    mvarUsePercent = False
    mvarShowCaption = True
    mvarMin = 0
    mvarMax = 100
    mvarValue = 0
    mvarPercent = 0
End Sub

Public Function GetVersion() As String
    GetVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function


'Private Sub subTestoVerticale(pcb As PictureBox, ByVal testo As String, ByVal x As Single, ByVal y As Single)
'
'    Dim f As LOGFONT
'    Dim hFont As Long, hFontOld As Long
'    f.lfEscapement = 900 'rotazione di 90 gradi
'    f.lfHeight = 14
'    f.lfCharSet = 1
'    f.lfQuality = 1
'    f.lfFaceName = "Verdana" & vbNullChar
'
'    hFont = CreateFontIndirect(f)
'    hFontOld = SelectObject(pcb.hdc, hFont)
'
'    TextOut pcb.hdc, x, y, testo, Len(testo)
'    'oppure dovrebbe funzionare anche se scrivi:
'    'pcb.CurrentX = x
'    'pcb.CurrentY = y
'    'pcb.Print testo
'
'    hFontOld = SelectObject(pcb.hdc, hFontOld)
'    DeleteObject hFont
'
'End Sub
'
Private Sub Class_Terminate()
    Set Picture = Nothing
End Sub

Public Sub GiveOfficeBorder()
    
    Dim lngRetVal As Long
    
    Picture.BorderStyle = 0
    Picture.Appearance = 0
    
    
    'Retrieve the current border style
    lngRetVal = GetWindowLong(Picture.hWnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
    'Apply the changes
    SetWindowLong Picture.hWnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos Picture.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
End Sub

Public Sub AboutBox()
    frmAbout.Show vbModal
End Sub
