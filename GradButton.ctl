VERSION 5.00
Begin VB.UserControl GradButton 
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   ScaleHeight     =   660
   ScaleWidth      =   1455
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   0
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "GradButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Control   : GradButton
' DateTime  : 4/25/2005 11:06
' Author    : Ken Foster
' Purpose   : Make Gradient Buttons
' Credits   : Middle-out gradient code by Matthew R. Usner
'---------------------------------------------------------------------------------------
Option Explicit

   ' Constants
   Const def_m_Caption = "GradButt"
   Const def_m_ColorOuter = vbBlack
   Const def_m_ColorMid = vbWhite
   Const def_m_ForeColor = vbBlack
   Const def_m_FontposX = 8
   Const def_m_FontposY = 3
   
   ' varibles
   Dim m_FontposX As Integer
   Dim m_FontposY As Integer
   Dim m_ForeColor As Long
   Dim m_ColorOuter As OLE_COLOR
   Dim m_ColorMid As OLE_COLOR
   Dim m_Caption As String
   Dim lcolor1 As Long
   Dim lcolor2 As Long
   
   ' events
   Event Click()
   Event MouseDown()
   Event MouseMove()
   Event Mouseup()
   
   ' Declarations
   Private Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
   Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
   Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
   Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
   Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
   Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Sub DrawGradient(ByVal hDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, _
   ByVal lCol1 As Long, ByVal lCol2 As Long)
   
   Dim yStart As Long
   Dim xEnd    As Long, yEnd   As Long
   Dim X1      As Long, Y1     As Long
   Dim X2      As Long, Y2     As Long
   Dim lRange  As Long
   Dim iQ      As Integer
   Dim lPtr    As Long, lInc   As Long
   Dim lCols() As Long, lCols2() As Long
   Dim hPO     As Long, hPN    As Long
   Dim r       As Long
   Dim x       As Long, xUp    As Long
   Dim B1(2)   As Byte, B2(2)  As Byte, b3(2) As Byte
   Dim p       As Single, ip   As Single
   Dim y As Long
   
   lInc = 1
   xEnd = (lWidth - 1) '/ 2
   yEnd = (lHeight - 1) '/ 2 ' /2 for top half only
   
   lRange = lHeight + yStart - 1
   
   X1 = IIf(iQ Mod 2, 0, xEnd)
   X2 = IIf(X1, -1, lWidth)
   '  -------------------------------------------------------------------
   '  Fill in the color array with the interpolated color values.
   '  -------------------------------------------------------------------
   ReDim lCols(lRange)
   ReDim lCols2(lRange)
   
   ' Get the r, g, b components of each color.
   RtlMoveMemory B1(0), lCol1, 3
   RtlMoveMemory B2(0), lCol2, 3
   RtlMoveMemory b3(0), 0, 3
   xUp = UBound(lCols)
   
   '        get the full color array in lCols2.
   For x = 0 To xUp
      ' Get the position and the 1 - position.
      p = x / xUp
      ip = 1 - p
      ' Interpolate the value at the current position.
      lCols2(x) = RGB(B1(0) * ip + B2(0) * p, B1(1) * ip + B2(1) * p, B1(2) * ip + B2(2) * p)
   Next x
   '        put the array in first half of lcols1
   y = 0
   For x = 0 To xUp Step 2
      lCols(y) = lCols2(x)
      y = y + 1
   Next x
   For x = xUp - 1 To 1 Step -2
      lCols(y) = lCols2(x)
      If y < xUp Then y = y + 1
   Next x
   
   For Y1 = -yStart To yEnd
      hPN = CreatePen(0, 1, lCols(lPtr))
      hPO = SelectObject(hDC, hPN)
      MoveTo hDC, X1, Y1, ByVal 0&
      LineTo hDC, X2, Y2
      r = SelectObject(hDC, hPO): r = DeleteObject(hPN)
      lPtr = lPtr + lInc
      Y2 = Y2 + 1
   Next Y1
   
End Sub

Public Sub UpdateGradient()
   Dim p As Integer
   Dim B As Integer
   
   lcolor1 = ColorOuter
   lcolor2 = ColorMid
   
   DrawGradient picButton.hDC, picButton.ScaleWidth, picButton.ScaleHeight, lcolor1, lcolor2
   picButton.AutoRedraw = True
   picButton.Font = Font
   picButton.ScaleMode = 3
   picButton.CurrentX = FontposX
   picButton.CurrentY = FontposY
   picButton.FontSize = Font.Size
   picButton.ForeColor = ForeColor
   picButton.FontBold = True
   picButton.Print Caption
   
End Sub

Private Sub picButton_Click()

   RaiseEvent Click
End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown
End Sub

Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent Mouseup
End Sub

Private Sub UserControl_Initialize()

   Caption = def_m_Caption
   ColorOuter = def_m_ColorOuter
   ColorMid = def_m_ColorMid
   ForeColor = def_m_ForeColor
   FontposX = def_m_FontposX
   FontposY = def_m_FontposY
   UpdateGradient
End Sub

Private Sub UserControl_Resize()
   picButton.Top = 0
   picButton.Left = 0
   picButton.Height = UserControl.Height
   picButton.Width = UserControl.Width
   UpdateGradient
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   m_Caption = PropBag.ReadProperty("Caption", def_m_Caption)
   Set Font = PropBag.ReadProperty("Font", Ambient.Font)
   m_ColorOuter = PropBag.ReadProperty("ColorOuter", m_ColorOuter)
   m_ColorMid = PropBag.ReadProperty("ColorMid", m_ColorMid)
   m_ForeColor = PropBag.ReadProperty("ForeColor", m_ForeColor)
   m_FontposX = PropBag.ReadProperty("FontposX", m_FontposX)
   m_FontposY = PropBag.ReadProperty("FontposY", m_FontposY)
   UpdateGradient
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   PropBag.WriteProperty "Caption", m_Caption, def_m_Caption
   PropBag.WriteProperty "Font", Font, Ambient.Font
   PropBag.WriteProperty "ColorOuter", m_ColorOuter, def_m_ColorOuter
   PropBag.WriteProperty "ColorMid", m_ColorMid, def_m_ColorMid
   PropBag.WriteProperty "ForeColor", m_ForeColor, def_m_ForeColor
   PropBag.WriteProperty "FontposX", m_FontposX, def_m_FontposX
   PropBag.WriteProperty "FontposY", m_FontposY, def_m_FontposY
End Sub

Public Property Get Caption() As String

   Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)

   m_Caption = NewCaption
   PropertyChanged "Caption"
   UpdateGradient
End Property

Public Property Get Font() As Font

   Set Font = picButton.Font
End Property

Public Property Set Font(NewFont As Font)

   Set picButton.Font = NewFont
   PropertyChanged ("Font")
   UpdateGradient
End Property

Public Property Get ColorOuter() As OLE_COLOR

   ColorOuter = m_ColorOuter
End Property

Public Property Let ColorOuter(NewColorOuter As OLE_COLOR)

   m_ColorOuter = NewColorOuter
   lcolor1 = m_ColorOuter
   PropertyChanged "ColorOuter"
   UpdateGradient
End Property

Public Property Get ColorMid() As OLE_COLOR

   ColorMid = m_ColorMid
End Property

Public Property Let ColorMid(NewColorMid As OLE_COLOR)

   m_ColorMid = NewColorMid
   lcolor2 = m_ColorMid
   PropertyChanged "ColorMid"
   UpdateGradient
End Property

Public Property Get ForeColor() As OLE_COLOR

   ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(NewForeColor As OLE_COLOR)

   m_ForeColor = NewForeColor
   PropertyChanged "ForeColor"
   UpdateGradient
End Property

Public Property Get FontposX() As Integer
   FontposX = m_FontposX
End Property

Public Property Let FontposX(NewFontposX As Integer)
   m_FontposX = NewFontposX
   PropertyChanged ("FontposX")
   UpdateGradient
End Property

Public Property Get FontposY() As Integer
   FontposY = m_FontposY
End Property

Public Property Let FontposY(NewFontposY As Integer)
   m_FontposY = NewFontposY
   PropertyChanged ("FontposY")
   UpdateGradient
End Property
