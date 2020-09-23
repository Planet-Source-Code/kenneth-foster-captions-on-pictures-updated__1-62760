VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00800080&
   Caption         =   "Add Captions to Pictures"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin Project1.GradButton cmdExit 
      Height          =   435
      Left            =   9495
      TabIndex        =   22
      Top             =   105
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOuter      =   255
      FontposX        =   20
      FontposY        =   6
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00800080&
      Caption         =   "Transparent Background"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   75
      TabIndex        =   20
      Top             =   4965
      Width           =   1215
   End
   Begin Project1.GradButton cmdDelete 
      Height          =   420
      Left            =   75
      TabIndex        =   19
      Top             =   5370
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   741
      Caption         =   "Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOuter      =   8421631
      FontposX        =   22
      FontposY        =   6
   End
   Begin Project1.GradButton cmdSaveJPG 
      Height          =   435
      Left            =   90
      TabIndex        =   16
      Top             =   945
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   767
      Caption         =   "Save Jpeg"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOuter      =   8421631
      FontposY        =   6
   End
   Begin Project1.GradButton cmdNextCaption 
      Height          =   405
      Left            =   60
      TabIndex        =   15
      Top             =   4530
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   714
      Caption         =   "New Caption"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOuter      =   8421504
      FontposX        =   4
      FontposY        =   6
   End
   Begin Project1.ThreeDText ThreeDText1 
      Height          =   735
      Left            =   4425
      TabIndex        =   11
      Top             =   -105
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   1296
      Caption         =   "Add Captions To Pictures"
      ColorS          =   12640511
      ColorE          =   8388736
      ColorF          =   8454143
      Direction       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.GradButton cmdFont 
      Height          =   420
      Left            =   90
      TabIndex        =   9
      Top             =   2430
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   741
      Caption         =   "Font Props"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOuter      =   8454143
      FontposX        =   12
      FontposY        =   6
   End
   Begin Project1.GradButton cmdFontColor 
      Height          =   405
      Left            =   90
      TabIndex        =   8
      Top             =   1980
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Font Color"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOuter      =   16744576
      FontposX        =   12
      FontposY        =   6
   End
   Begin Project1.GradButton cmdClear 
      Height          =   390
      Left            =   90
      TabIndex        =   7
      Top             =   1410
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   688
      Caption         =   "Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOuter      =   8438015
      FontposX        =   25
      FontposY        =   6
   End
   Begin Project1.GradButton cmdSave 
      Height          =   375
      Left            =   90
      TabIndex        =   6
      Top             =   540
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   661
      Caption         =   "Save Bitmap"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOuter      =   8454016
      FontposX        =   5
      FontposY        =   6
   End
   Begin Project1.GradButton cmdLoad 
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   661
      Caption         =   "Load"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorOuter      =   16744703
      FontposX        =   25
      FontposY        =   6
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   10665
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCaption 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1470
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   75
      Width           =   1890
   End
   Begin VB.PictureBox picS 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6795
      Left            =   1470
      ScaleHeight     =   6735
      ScaleWidth      =   9540
      TabIndex        =   1
      Top             =   660
      Width           =   9600
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   4155
         TabIndex        =   17
         Top             =   3480
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.PictureBox picD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6735
      Left            =   1485
      ScaleHeight     =   6675
      ScaleWidth      =   9435
      TabIndex        =   18
      Top             =   675
      Width           =   9495
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   2595
      Left            =   30
      Top             =   3885
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Double Click Caption then Delete Button"
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   240
      TabIndex        =   21
      Top             =   5820
      Width           =   1065
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by     Ken Foster          @ 2005"
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   135
      TabIndex        =   14
      Top             =   6900
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   90
      TabIndex        =   13
      Top             =   6615
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Click and Drag to position caption on picture."
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   90
      TabIndex        =   12
      Top             =   3885
      Width           =   1275
   End
   Begin VB.Label lblFontColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Font Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   90
      TabIndex        =   10
      Top             =   3525
      Width           =   1260
   End
   Begin VB.Label lblFontname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   3210
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Font Size"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   525
      TabIndex        =   4
      Top             =   2910
      Width           =   720
   End
   Begin VB.Label lblFontsize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   2910
      Width           =   315
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1890
      Left            =   45
      Top             =   1935
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1800
      Left            =   45
      Top             =   75
      Width           =   1350
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************
'*
'*         Name of Program: Add Captions to Pictures
'*                  Author: Ken Foster
'*                 Version: 1.0.0
'*                    Date: October 01,2005
'*                    Time: 09:55 PM
'*         No Copyrights claimed - Use as you like
'*
'***************************************************
'   Some features
'   1. Save as bitmap or jpeg
'   2. Uses my Gradient button control
'   3. Uses my ThreeD Text control
'   4. Pre-Cursor positioning
'   5. Multi-captions
'***************** Table of Procedures *************
'   Private Sub Form_Load
'   Private Sub Form_MouseMove
'   Private Sub Form_Resize
'   Private Sub Form_Unload
'   Private Sub cmdDelete_Click
'   Private Sub cmdDelelt_MouseMove
'   Private Sub cmdExit_Click
'   Private Sub cmdFont_Click
'   Private Sub cmdFont_MouseMove
'   Private sun cmdFont_GotFocus
'   Private Sub cmdFontColor_Click
'   Private Sub cmdFontColor_MouseMove
'   Private Sub cmdFontColor_GotFocus
'   Private Sub cmdSave_Click
'   Private Sub cmdSave_MouseMove
'   Private Sub cmdSaveJPG
'   Private Sub cmoSaveJPG_MouseMove
'   Private Sub cmdLoad_Click
'   Private Sub cmdLoad_MouseMove
'   Private Sub cmdClear_Click
'   Private Sub cmdClear_MouseMove
'   Private Sub cmdNextCaption
'   Private Sub cmdNextCaption_MouseMove
'   Private Sub Label1_MouseDown
'   Private Sub Label1_MouseMove
'   Private Sub picS_Change
'   Private Sub picS_Mousemove
'   Private Sub txtCaption_Change
'   Private Sub txtCaption_MouseMove
'   Private Sub cpyLabel
'   Private Sub focus
'   Private Function CheckExt
'   Private Function SaveJPEG
'***************** End of Table ********************
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Private OldX As Integer
Private OldY As Integer
Dim ind As Integer                                              ' current caption label
Dim indel As Integer                                            ' caption label to be deleted
Dim NSD As Boolean                                              ' Pic not saved

Private Sub Form_Load()
   
   ind = 0
   dlgCommon.CancelError = True                                  'catches errors that occur when the user hits cancel
   dlgCommon.FontName = picS.FontName
   lblFontname.Caption = dlgCommon.FontName
   lblFontsize.Caption = dlgCommon.FontSize
   NSD = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Label4.Caption = ""
End Sub

Private Sub Form_Resize()
   If picS.Width + 1700 < 11340 Then                             ' make form wider than picture
      Form1.Width = 11340
   Else
      Form1.Width = picS.Width + 1700
   End If
   
   If picS.Height + 1300 < 8070 Then                             ' make form taller than picture
      Form1.Height = 8070
   Else
      Form1.Height = picS.Height + 1300
   End If
   picD.Width = picS.Width                                       ' make hidden Destination pic same  as Source pic
   picD.Height = picS.Height
   picD.Top = picS.Top
   picD.Left = picS.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub cmdDelete_Click()
   NSD = True
   lblCap(indel).Visible = False                                ' hide label
End Sub

Private Sub cmdDelete_MouseMove()
   Label4.Caption = "Delete"                                    ' show button name in label box
End Sub

Private Sub cmdDelete_GotFocus()
   Call focus
End Sub

Private Sub cmdExit_Click()
   Dim iresponse As String
   
   If NSD = True Then
       iresponse = MsgBox("Picture not saved.Do you want to save it ?", vbYesNo, "File Exists")
       If iresponse = vbYes Then
          Exit Sub
       Else
          Unload Me
       End If
   End If
   If NSD = False Then Unload Me
End Sub

Private Sub cmdFont_Click()
On Error GoTo FontErr                                            'catches error when user hits cancel
   'loads the fonts
   dlgCommon.Flags = cdlCFScreenFonts
   '================
   dlgCommon.ShowFont
   '================
   'changes the settings according to the commondialog changes
   txtCaption.FontName = dlgCommon.FontName
   txtCaption.FontBold = dlgCommon.FontBold
   txtCaption.FontItalic = dlgCommon.FontItalic
   txtCaption.FontSize = dlgCommon.FontSize
   
   lblCap(ind).Font = dlgCommon.FontName
   lblFontsize = dlgCommon.FontSize
   lblFontname = dlgCommon.FontName
   cmdNextCaption.SetFocus
   
FontErr:
   Exit Sub
End Sub

Private Sub cmdFont_MouseMove()
   Label4.Caption = "Font Props"
End Sub

Private Sub cmdFont_GotFocus()
   Call focus                                                    ' move cursor to Font button
End Sub

Private Sub cmdFontColor_Click()
On Error GoTo ColorErr                                           'catches error when user hits cancel
   'selects black when loads
   dlgCommon.Flags = cdlCCRGBInit
   '==================
   dlgCommon.ShowColor
   '==================
   If dlgCommon.Color <> vbBlack Then                            ' if colorlabel backgound is black make font white and visa versa
      lblFontColor.ForeColor = vbBlack
      lblFontColor.BackColor = dlgCommon.Color
   Else
      lblFontColor.ForeColor = vbWhite
      lblFontColor.BackColor = dlgCommon.Color
   End If
   
   If dlgCommon.Color = vbWhite Then                             ' if Caption textbox background is black make font white and visa versa
      txtCaption.BackColor = vbBlack
      txtCaption.ForeColor = vbWhite
   Else
      txtCaption.BackColor = vbWhite
      txtCaption.ForeColor = vbBlack
   End If
   cmdFont.SetFocus                                              ' used to position cursor to next button
ColorErr:
   Exit Sub
End Sub

Private Sub cmdFontColor_MouseMove()
   Label4.Caption = "Font Color"
End Sub

Private Sub cmdFontColor_GotFocus()
   Call focus                                                    ' used to position cursor to FontColor Button
End Sub

Private Sub cmdSave_Click() ' Save as bitmap
Dim iresponse As String
Dim Fname As String

On Error GoTo SaveErr

   picD.Picture = picD.Image                                     ' copy picS to picD so we can save it
   BitBlt picD.hDC, 0, 0, picS.ScaleWidth, picS.ScaleHeight, picS.hDC, 0, 0, vbSrcCopy
   
   dlgCommon.DialogTitle = "Save As Bitmap"
   dlgCommon.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist ' sets flags to overwrite file and pathmustexist
   dlgCommon.Filter = "Bitmap (*.bmp)|*.bmp"                     ' sets the file type
   '=================
   dlgCommon.ShowSave
   '=================
   
   dlgCommon.Filename = CheckExt(dlgCommon.Filename, ".bmp")     ' get extension if not present or not correct

   ' give dialog window time to close properly
   SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    If MsgBox("Formatting...Please Wait", vbOKOnly, "Self Closing Message Box") = vbOKOnly Then
    End If
   
   SavePicture picD.Image, dlgCommon.Filename                    ' save picture as bitmap
   MsgBox "Bitmap saved in " & dlgCommon.Filename
   NSD = False
SaveErr:
   Exit Sub
End Sub

Private Sub cmdSave_MouseMove()
   Label4.Caption = "Save Bitmap"
End Sub

Private Sub cmdSaveJPG_Click() ' Save as Jpeg
Dim iresponse As String
Dim Fname As String
   On Error GoTo JpgErr
   picD.Picture = picD.Image                                     ' copy picS to picD so we can save it
   BitBlt picD.hDC, 0, 0, picS.ScaleWidth, picS.ScaleHeight, picS.hDC, 0, 0, vbSrcCopy
   
   dlgCommon.DialogTitle = "Save As Jpeg"
   dlgCommon.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist ' sets flags to overwrite file and pathmustexist
   dlgCommon.Filter = "Jpeg (*.jpg)|*.jpg"                     ' sets the file type
   '===================
   dlgCommon.ShowSave
   '===================
   dlgCommon.Filename = CheckExt(dlgCommon.Filename, ".jpg")
   
    ' give dialog window time to close properly
    SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
    If MsgBox("Formatting...Please Wait", vbOKOnly, "Self Closing Message Box") = vbOKOnly Then
    End If

   If SaveJPEG(dlgCommon.Filename, picD, True, 90) = True Then   ' save pic as Jpeg
      MsgBox "Jpeg saved in folder " & dlgCommon.Filename
   End If
   NSD = False
JpgErr:
   Exit Sub
End Sub

Private Sub cmdSaveJPG_MouseMove()
   Label4.Caption = "Save Jpeg"
End Sub

Private Sub cmdLoad_Click()
On Error GoTo LoadErr
   cmdClear_Click
   dlgCommon.DialogTitle = "Load an Image"
   ' sets the file type
   dlgCommon.Filter = "All Files (*.*)|*.*|Bitmap (*.bmp)|*.bmp|JPeg (*.jpg)|*.jpg|Gif (*.gif)|*.gif"
   '=================
   dlgCommon.ShowOpen
   '=================
   picS.Picture = LoadPicture(dlgCommon.Filename, , , 0, 0)           ' load picture
   lblCap(ind).Top = picS.Height / 2                                  ' move label back to center of picture
   lblCap(ind).Left = picS.Width / 4                                  ' move label back to center of picture
   Form1.Caption = "Add Captions to Pictures... " & dlgCommon.Filename
   NSD = True
   cmdFontColor.SetFocus
LoadErr:
   Exit Sub
End Sub

Private Sub cmdLoad_MouseMove()
   Label4.Caption = "Load Picture"
End Sub

Private Sub cmdLoad_GotFocus()
    Call focus                                                   ' move cursor to Load Button
End Sub

Private Sub cmdClear_Click()
Dim XP As Integer
   picS.Picture = LoadPicture                                    ' clear picture
   picD.Picture = LoadPicture
   
   For XP = 0 To ind
     lblCap(XP).Caption = ""                                      ' clear label
     lblCap(XP).Visible = False
   Next XP
   
    txtCaption.Text = ""
End Sub

Private Sub cmdClear_MouseMove()
   Label4.Caption = "Clear"
End Sub

Private Sub cmdNextCaption_Click()

   ind = ind + 1
  
   Load lblCap(ind)                                              ' load new label
    lblCap(ind).Top = picS.Height / 2
    lblCap(ind).Left = picS.Width / 4
    lblCap(ind).Visible = True
    lblCap(ind).ForeColor = dlgCommon.Color
    lblCap(ind).FontSize = dlgCommon.FontSize
    
    If Check1.Value = 1 Then                                     ' transparent background
      lblCap(ind).BackStyle = 0
      lblCap(ind).BorderStyle = 0
    Else
      lblCap(ind).BackStyle = 1
      lblCap(ind).BorderStyle = 1
    End If
    NSD = True
    txtCaption.Text = ""
    txtCaption.SetFocus
End Sub

Private Sub cmdNextCaption_MouseMove()
   Label4.Caption = "New Caption"
End Sub

Private Sub cmdNextCaption_GotFocus()
   Call focus
End Sub

Private Sub lblCap_DblClick(Index As Integer)
   If lblCap(Index).ForeColor <> vbRed Then                     ' change label forecolor to indicate a delete
      lblCap(Index).ForeColor = vbRed
   Else
      lblCap(Index).ForeColor = dlgCommon.Color
   End If
   indel = Index                                                ' set to the caption we are going to delete
   cmdDelete.SetFocus
End Sub

Private Sub lblCap_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   OldX = x                                                      ' used to move label
   OldY = y                                                      ' used to move label
End Sub

Private Sub lblCap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 Then                                            ' if left mousebutton pressed move label
      lblCap(Index).Left = lblCap(Index).Left + (x - OldX)
      lblCap(Index).Top = lblCap(Index).Top + (y - OldY)
      NSD = True
   End If
End Sub

Private Sub lblFontColor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Label4.Caption = "Font Color"
End Sub

Private Sub lblFontname_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Label4.Caption = "Font Name"
End Sub

Private Sub picS_Change()
   Form_Resize                                                   ' resize form to picture
End Sub

Private Sub picS_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Label4.Caption = ""
End Sub

Private Sub txtCaption_Change()
   
   'txtCaption.FontBold = dlgCommon.FontBold
   'txtCaption.FontSize = dlgCommon.FontSize
   'txtCaption.ForeColor = dlgCommon.Color
   
   lblCap(ind).ForeColor = dlgCommon.Color
   lblCap(ind).FontBold = dlgCommon.FontBold
   lblCap(ind).FontSize = dlgCommon.FontSize
   lblCap(ind).Caption = txtCaption.Text
End Sub

Private Sub txtCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Label4.Caption = "Enter Caption"
End Sub

Private Sub focus()                                               ' moves cursor to selected control with focus
Dim x&                                                            ' put "Call focus" in the GotFocus procedure of control
Dim y&
Dim a&
' this is not my code..it is from PSC
If Me.BorderStyle = 0 Then
   x& = Me.ActiveControl.Left \ Screen.TwipsPerPixelX _
   + ((Me.ActiveControl.Width / 2) / Screen.TwipsPerPixelX) _
   + (Me.Left / Screen.TwipsPerPixelX)
   y& = Me.ActiveControl.Top \ Screen.TwipsPerPixelY _
   + ((Me.ActiveControl.Height / 2) / Screen.TwipsPerPixelY) _
   + (Me.Top / Screen.TwipsPerPixelY)
Else
   x& = Me.ActiveControl.Left \ Screen.TwipsPerPixelX + ((Me.ActiveControl.Width / 2 + 60) / Screen.TwipsPerPixelX) + (Me.Left / Screen.TwipsPerPixelX)
   ' "+ 60" is for the border"
   y& = Me.ActiveControl.Top \ Screen.TwipsPerPixelY + ((Me.ActiveControl.Height / 2 + 360) / Screen.TwipsPerPixelY) + (Me.Top / Screen.TwipsPerPixelY)
   ' "+ 360 " is for the border and the tittle bar"
End If
a& = SetCursorPos(x&, y&)
End Sub

Private Function CheckExt(Filename As String, ext As String)
    Dim stg1 As String
    Dim stg2 As String
    
    stg1 = Right$(Filename, 4)                                      ' get extension
    stg2 = Left$(Filename, Len(Filename) - 4)                       ' get filename without extension
    
    If InStr(Filename, ".") = False Then                            ' if no extension present
       CheckExt = Filename & ext                                    ' add extension to filename
    Else
       CheckExt = stg2 & ext                                        ' Makes sure we have correct extension
    End If
      
End Function

Private Function SaveJPEG(ByVal Filename As String, Pic As PictureBox, Optional ByVal Overwrite As Boolean = True, Optional ByVal Quality As Byte = 90) As Boolean
    Dim JPEGclass As cJpeg
    Dim m_Picture As IPictureDisp
    Dim m_DC As Long
    Dim m_Millimeter As Single
    m_Millimeter = ScaleX(100, vbPixels, vbMillimeters)
    Set m_Picture = Pic
    m_DC = Pic.hDC
    'this is not my code....from PSC
    'initialize class
    Set JPEGclass = New cJpeg
    'check there is image to save and the filename string is not empty
    If m_DC <> 0 And LenB(Filename) > 0 Then
        'check for valid quality
        If Quality < 1 Then Quality = 1
        If Quality > 100 Then Quality = 100
        'set quality
        JPEGclass.Quality = Quality
        'save in full color
        JPEGclass.SetSamplingFrequencies 1, 1, 1, 1, 1, 1
        'copy image from hDC
        If JPEGclass.SampleHDC(m_DC, CLng(m_Picture.Width / m_Millimeter), CLng(m_Picture.Height / m_Millimeter)) = 0 Then
            'if overwrite is set and file exists, delete the file
            If Overwrite And LenB(Dir$(Filename)) > 0 Then Kill Filename
            'save file and return True if success
            SaveJPEG = JPEGclass.SaveFile(Filename) = 0
        End If
    End If
    'clear memory
    Set JPEGclass = Nothing
End Function
