VERSION 5.00
Begin VB.UserControl FrameXP 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "FrameXP.ctx":0000
End
Attribute VB_Name = "FrameXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'XP style frame usercontrol
'By Jim K on Dec, 05

'This usercontrol is based on the work of: Juan Carlos San Román
'Find it here: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=63739&lngWId=1#FeedbackId_285296

'Credits to: Juan Carlos San Román for the
'PaintFrame() Sub. And thanks for the idea
'to make this usercontrol

'The sub is mainly as Juan posted it, but a few
'changes/addings are made. See comments in sub()
'Additionally to his original code, you can define
'BorderColor, CaptionColor, Fonts and Ellipse size
'on the corners.


'API declarations
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private lpp As POINTAPI

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_LEFT = &H0

'UC Defaults
Const mDefCaptionColor = vbBlack
Const mDefBorderColor = vbBlue
Const mDefBackColor = vbButtonFace
Const mDefCaption = ""
Const mDefCornerCurveSize = 8 'Max is set to 30. Also, I see no point in
                              'changing ellipse height and width separately.
                              'If you want the mentioned propertys, then
                              'change/add code for it yourself

'Property holders
Private mCaptionCol As OLE_COLOR
Private mBorderCol As OLE_COLOR
Private mBackCol As OLE_COLOR
Private mCaption As String
Private mCCSize As Integer

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackCol
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
    mBackCol = NewBackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = mBackCol
    PaintFrame
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal NewFont As Font)
    Set UserControl.Font = NewFont
    PropertyChanged "Font"
    PaintFrame
End Property

Public Property Get CornerCurveSize() As Integer
    CornerCurveSize = mCCSize
End Property

Public Property Let CornerCurveSize(ByVal NewCornerCurveSize As Integer)
    mCCSize = NewCornerCurveSize
    PropertyChanged "CornerCurveSize"
    If mCCSize < 0 Then
        MsgBox "Below 0 is not allowed", vbInformation, "Info"
        mCCSize = 0
    ElseIf mCCSize > 30 Then
        MsgBox "Max is set to 30." & vbCrLf & _
        "You can change that in the source.", vbInformation, "Info"
        mCCSize = 30
    End If
    PaintFrame
End Property

Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    mCaption = NewCaption
    PropertyChanged "Caption"
    PaintFrame
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mBorderCol
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
    mBorderCol = NewBorderColor
    PropertyChanged "BorderColor"
    PaintFrame
End Property

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = mCaptionCol
End Property

Public Property Let CaptionColor(ByVal NewCaptionColor As OLE_COLOR)
    mCaptionCol = NewCaptionColor
    PropertyChanged "CaptionColor"
    PaintFrame
End Property

Private Sub UserControl_InitProperties()
    CaptionColor = mDefCaptionColor
    BorderColor = mDefBorderColor
    BackColor = mDefBackColor
    Caption = Ambient.DisplayName
    CornerCurveSize = mDefCornerCurveSize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mCaption = PropBag.ReadProperty("Caption", mDefCaption)
    mBorderCol = PropBag.ReadProperty("BorderColor", mDefBorderColor)
    mBackCol = PropBag.ReadProperty("BackColor", mDefBackColor)
    mCaptionCol = PropBag.ReadProperty("CaptionColor", mDefCaptionColor)
    mCCSize = PropBag.ReadProperty("CornerCurveSize", mDefCornerCurveSize)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    PaintFrame
End Sub

Private Sub UserControl_Resize()
    PaintFrame
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", mCaption, mDefCaption
    PropBag.WriteProperty "BorderColor", mBorderCol, mDefBorderColor
    PropBag.WriteProperty "CaptionColor", mCaptionCol, mDefCaptionColor
    PropBag.WriteProperty "BackColor", mBackCol, mDefBackColor
    PropBag.WriteProperty "CornerCurveSize", mCCSize, mDefCornerCurveSize
    PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
End Sub

Private Sub PaintFrame()
    
    Dim p_LenghOfCaption As Long, p_HeightOfCaption As Long
    Dim p_indentation As Long, p_offset As Long, R As RECT, space As Long
    Dim L As Long 'Added for Caption Left for more accurate
    'space between line left of text and text/text and line right of text

    UserControl.AutoRedraw = True  'Added so user don't have to
    'change this property for all pics he/she use as frames. Counts
    'only for Juans posting. That would be (Pic.AutoRedraw)
    UserControl.ScaleMode = vbPixels
    UserControl.Picture = LoadPicture()
    p_LenghOfCaption = UserControl.TextWidth(mCaption)
    p_HeightOfCaption = UserControl.TextHeight(mCaption)
    p_offset = 5
    p_indentation = 12
    space = 4
    L = 15
    
    'define back color
    UserControl.BackColor = mBackCol
    
    'define border color
    UserControl.ForeColor = mBorderCol
    
    'paint rounded rectangle
    RoundRect UserControl.hDC, p_offset, p_offset + p_HeightOfCaption / 2, UserControl.ScaleWidth - p_offset, UserControl.ScaleHeight - p_offset, mCCSize, mCCSize
    
    'define caption rectangle
    SetRect R, p_offset + p_indentation + space, p_offset, p_offset + p_indentation + p_LenghOfCaption + 2 * space, p_offset + p_HeightOfCaption
    
    'Set line color
    UserControl.ForeColor = UserControl.BackColor
    MoveToEx UserControl.hDC, p_offset + p_indentation + space, p_offset + p_HeightOfCaption / 2, lpp
    LineTo UserControl.hDC, p_offset + p_indentation + p_LenghOfCaption + 2 * space, p_offset + p_HeightOfCaption / 2
    
    'Set caption color
    UserControl.ForeColor = mCaptionCol
    
    'Draw frame caption
    DrawTextEx UserControl.hDC, mCaption, Len(mCaption), R, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER + L, ByVal 0&

    UserControl.ScaleMode = vbTwips
    UserControl.Picture = UserControl.Image
    
End Sub

