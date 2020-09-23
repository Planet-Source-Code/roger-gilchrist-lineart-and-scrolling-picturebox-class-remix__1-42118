VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Line Art And Edge Detection by Arivnder Sehmi (Modifed by Roger Gilchrist; thanks to Min Thant Sin for scrolling code)"
   ClientHeight    =   9780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13725
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   13725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      TabIndex        =   73
      Top             =   9000
      Width           =   735
   End
   Begin VB.CommandButton ContiansPictureOverRide 
      Caption         =   "Force ContainsImage"
      Height          =   495
      Left            =   6240
      TabIndex        =   48
      ToolTipText     =   "Because the test ContainsImage can fail if an image is mostly blank, this allows you to over-ride the test"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Paste2Source 
      Caption         =   "<------------ Paste to Source <------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   15
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton BtnGreyScale 
      Caption         =   "Grey Scale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12060
      TabIndex        =   13
      Top             =   7440
      Width           =   1500
   End
   Begin VB.CommandButton Load 
      Caption         =   "Load A Colour\Greyscale Image"
      Height          =   420
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   2715
   End
   Begin VB.CommandButton CancelDraw 
      Cancel          =   -1  'True
      Caption         =   "Cancel Draw"
      Enabled         =   0   'False
      Height          =   390
      Left            =   12060
      TabIndex        =   14
      Top             =   7920
      Width           =   1500
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save Line Art Image"
      Height          =   420
      Left            =   7560
      TabIndex        =   22
      Top             =   6000
      Width           =   2715
   End
   Begin VB.CommandButton StartEdgeDetect 
      Caption         =   "Edge Detection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12060
      TabIndex        =   12
      Top             =   6960
      Width           =   1500
   End
   Begin VB.Frame FramOptions 
      Caption         =   "Options:"
      Enabled         =   0   'False
      Height          =   3135
      Left            =   0
      TabIndex        =   21
      Top             =   6480
      Width           =   11895
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   11655
         TabIndex        =   7
         Top             =   180
         Width           =   11655
         Begin VB.Frame Frame8 
            Caption         =   "Experimental"
            Height          =   2775
            Left            =   0
            TabIndex        =   49
            ToolTipText     =   "Stuff in this box is not fully developed"
            Top             =   0
            Width           =   3735
            Begin VB.PictureBox Picture3 
               BorderStyle     =   0  'None
               Height          =   2415
               Left            =   120
               ScaleHeight     =   2415
               ScaleWidth      =   3495
               TabIndex        =   50
               Top             =   240
               Width           =   3495
               Begin VB.CommandButton Command2 
                  Caption         =   "Darken 2"
                  Height          =   255
                  Index           =   18
                  Left            =   1800
                  TabIndex        =   72
                  Top             =   2160
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Lighten"
                  Height          =   255
                  Index           =   17
                  Left            =   1800
                  TabIndex        =   71
                  Top             =   1920
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Negative4"
                  Height          =   255
                  Index           =   16
                  Left            =   240
                  TabIndex        =   70
                  Top             =   2160
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Negative3"
                  Height          =   255
                  Index           =   15
                  Left            =   240
                  TabIndex        =   69
                  Top             =   1920
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Negative2"
                  Height          =   255
                  Index           =   14
                  Left            =   240
                  TabIndex        =   68
                  Top             =   1680
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Negative1"
                  Height          =   255
                  Index           =   13
                  Left            =   240
                  TabIndex        =   67
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Aqua(RG) Filter"
                  Height          =   255
                  Index           =   10
                  Left            =   1800
                  TabIndex        =   66
                  Top             =   1680
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Purple(RB) Filter"
                  Height          =   255
                  Index           =   11
                  Left            =   1800
                  TabIndex        =   65
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Yellow(BG) Filter"
                  Height          =   255
                  Index           =   12
                  Left            =   1800
                  TabIndex        =   64
                  Top             =   1200
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Blue Filter"
                  Height          =   255
                  Index           =   9
                  Left            =   1800
                  TabIndex        =   63
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Green Filter"
                  Height          =   255
                  Index           =   8
                  Left            =   1800
                  TabIndex        =   62
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Red Filter"
                  Height          =   255
                  Index           =   7
                  Left            =   1800
                  TabIndex        =   61
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Brighten"
                  Height          =   255
                  Index           =   5
                  Left            =   1800
                  TabIndex        =   57
                  Top             =   0
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Darken"
                  Height          =   255
                  Index           =   6
                  Left            =   1800
                  TabIndex        =   56
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Diffuse"
                  Height          =   255
                  Index           =   4
                  Left            =   240
                  TabIndex        =   55
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Smooth"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   54
                  Top             =   0
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Sharp"
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   53
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Sharp 2"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   52
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Neon Crayon"
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   51
                  Top             =   720
                  Width           =   1335
               End
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Grey Scaling"
            Height          =   2055
            Left            =   7920
            TabIndex        =   41
            Top             =   0
            Width           =   1695
            Begin ComCtl2.UpDown UpDown1 
               Height          =   285
               Left            =   1320
               TabIndex        =   60
               Top             =   1080
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   327681
               AutoBuddy       =   -1  'True
               BuddyControl    =   "Text1"
               BuddyDispid     =   196623
               OrigLeft        =   1320
               OrigTop         =   1080
               OrigRight       =   1575
               OrigBottom      =   1335
               Max             =   5
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   1080
               TabIndex        =   59
               Text            =   "0"
               Top             =   1080
               Width           =   255
            End
            Begin VB.CheckBox huh 
               Alignment       =   1  'Right Justify
               Caption         =   "Huh?!?"
               Height          =   315
               Left            =   120
               TabIndex        =   43
               ToolTipText     =   $"MainFrm.frx":1472
               Top             =   1320
               Width           =   1455
            End
            Begin VB.CheckBox IgnoreTol 
               Alignment       =   1  'Right Justify
               Caption         =   "Ignore Tolerance"
               Height          =   435
               Left            =   120
               TabIndex        =   44
               Top             =   1560
               Width           =   1455
            End
            Begin VB.HScrollBar greys 
               Height          =   255
               LargeChange     =   25
               Left            =   120
               Max             =   255
               Min             =   1
               TabIndex        =   42
               Top             =   480
               Value           =   255
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "Grey Maths"
               Height          =   375
               Left            =   120
               TabIndex        =   58
               ToolTipText     =   "Use various different ways of calculating Grey scaling"
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label grys 
               BackStyle       =   0  'Transparent
               Caption         =   "255"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   1080
               TabIndex        =   46
               Top             =   240
               Width           =   405
            End
            Begin VB.Label Label4 
               Caption         =   "Greys(1-255):"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Edge Detection "
            Height          =   2055
            Left            =   5880
            TabIndex        =   25
            Top             =   0
            Width           =   1935
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "Change Tolerence while drawing"
               Height          =   435
               Left            =   120
               TabIndex        =   38
               ToolTipText     =   "Faster if you don't use this!"
               Top             =   1560
               Width           =   1650
            End
            Begin VB.CheckBox chkPixel 
               Caption         =   "Check1"
               Height          =   200
               Index           =   0
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   200
            End
            Begin VB.CheckBox chkPixel 
               Caption         =   "Check1"
               Height          =   200
               Index           =   1
               Left            =   360
               TabIndex        =   31
               Top             =   240
               Width           =   200
            End
            Begin VB.CheckBox chkPixel 
               Caption         =   "Check1"
               Height          =   200
               Index           =   2
               Left            =   600
               TabIndex        =   32
               Top             =   240
               Width           =   200
            End
            Begin VB.CheckBox chkPixel 
               Caption         =   "Check1"
               Height          =   200
               Index           =   3
               Left            =   600
               TabIndex        =   33
               Top             =   480
               Width           =   200
            End
            Begin VB.CheckBox chkPixel 
               Caption         =   "Check1"
               Height          =   200
               Index           =   4
               Left            =   600
               TabIndex        =   34
               Top             =   720
               Width           =   200
            End
            Begin VB.CheckBox chkPixel 
               Caption         =   "Check1"
               Height          =   200
               Index           =   5
               Left            =   360
               TabIndex        =   35
               Top             =   720
               Width           =   200
            End
            Begin VB.CheckBox chkPixel 
               Caption         =   "Check1"
               Height          =   200
               Index           =   6
               Left            =   120
               TabIndex        =   36
               Top             =   720
               Width           =   200
            End
            Begin VB.CheckBox chkPixel 
               Caption         =   "Check1"
               Height          =   200
               Index           =   7
               Left            =   120
               TabIndex        =   37
               Top             =   480
               Width           =   200
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "MainFrm.frx":151A
               Left            =   960
               List            =   "MainFrm.frx":151C
               TabIndex        =   29
               ToolTipText     =   "Test Any | All pixels checked"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "General "
            Height          =   1575
            Left            =   9720
            TabIndex        =   24
            Top             =   0
            Width           =   1935
            Begin VB.HScrollBar Tolerance 
               Height          =   255
               LargeChange     =   25
               Left            =   120
               Max             =   255
               TabIndex        =   19
               Top             =   480
               Width           =   1695
            End
            Begin VB.CheckBox Invert 
               Alignment       =   1  'Right Justify
               Caption         =   "Invert Image"
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   1200
               Width           =   1650
            End
            Begin VB.Label Tol 
               BackStyle       =   0  'Transparent
               Caption         =   "157"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   1320
               TabIndex        =   28
               Top             =   240
               Width           =   405
            End
            Begin VB.Label Label1 
               Caption         =   "Tolerance(0-255):"
               Height          =   180
               Left            =   45
               TabIndex        =   27
               Top             =   240
               Width           =   1305
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Line Art "
            Height          =   1575
            Left            =   3840
            TabIndex        =   23
            Top             =   0
            Width           =   1935
            Begin VB.CheckBox uBlue 
               Caption         =   "Blue Values"
               Height          =   240
               Left            =   120
               TabIndex        =   18
               Top             =   1080
               Value           =   1  'Checked
               Width           =   1410
            End
            Begin VB.CheckBox uGreen 
               Caption         =   "Green Values"
               Height          =   240
               Left            =   120
               TabIndex        =   17
               Top             =   840
               Value           =   1  'Checked
               Width           =   1410
            End
            Begin VB.CheckBox uRed 
               Caption         =   "Red Values"
               Height          =   240
               Left            =   120
               TabIndex        =   16
               Top             =   600
               Value           =   1  'Checked
               Width           =   1410
            End
            Begin VB.Label Label2 
               Caption         =   "Use when Grey Scaling:"
               Height          =   240
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   1770
            End
         End
      End
   End
   Begin VB.CommandButton StartLineArt 
      Caption         =   "Draw Line Art"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12060
      TabIndex        =   11
      Top             =   6480
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Caption         =   "clsScrollPicture captions me"
      Height          =   5745
      Left            =   7560
      TabIndex        =   5
      Top             =   0
      Width           =   5520
      Begin VB.VScrollBar DScrollV 
         Enabled         =   0   'False
         Height          =   1785
         LargeChange     =   500
         Left            =   1800
         SmallChange     =   150
         TabIndex        =   8
         Top             =   1680
         Width           =   150
      End
      Begin VB.HScrollBar DScrollH 
         Enabled         =   0   'False
         Height          =   150
         LargeChange     =   500
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   2835
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   3465
         Left            =   120
         ScaleHeight     =   3465
         ScaleWidth      =   3675
         TabIndex        =   2
         Top             =   240
         Width           =   3675
         Begin VB.PictureBox Dest 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1545
            Left            =   960
            ScaleHeight     =   1545
            ScaleWidth      =   1515
            TabIndex        =   6
            Top             =   360
            Width           =   1515
         End
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"MainFrm.frx":151E
         Height          =   1815
         Left            =   120
         TabIndex        =   39
         Top             =   3840
         Visible         =   0   'False
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6000
      Left            =   120
      TabIndex        =   0
      Tag             =   "V-hi"
      Top             =   0
      Width           =   6000
      Begin VB.HScrollBar SScrollH 
         Enabled         =   0   'False
         Height          =   150
         LargeChange     =   500
         Left            =   0
         TabIndex        =   3
         Top             =   4080
         Width           =   5595
      End
      Begin VB.VScrollBar SScrollV 
         Enabled         =   0   'False
         Height          =   5505
         LargeChange     =   500
         Left            =   4560
         SmallChange     =   150
         TabIndex        =   4
         Top             =   -240
         Width           =   150
      End
      Begin VB.PictureBox SourceContainer 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   3705
         Left            =   120
         ScaleHeight     =   3705
         ScaleWidth      =   4275
         TabIndex        =   47
         Top             =   600
         Width           =   4275
         Begin VB.PictureBox Source 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   3210
            Left            =   720
            ScaleHeight     =   3210
            ScaleWidth      =   3255
            TabIndex        =   1
            Top             =   120
            Width           =   3255
         End
      End
   End
   Begin VB.Label PercentDone 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   12060
      TabIndex        =   40
      Top             =   6120
      Width           =   1500
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------'
' Line Art Creation, And Edge Dection By Arivnder Sehmi. '
' Arvinder@Sehmi.org.uk                                  '
' September 23th 2000                                    '
'                                                        '
' Modifications by Roger Gilchrist                       '
' rojagilkrist@hotmail.com                               '
' December 17 2002                                       '
'--------------------------------------------------------'

Option Explicit
Private OSF As New ClsOpenSave
Private LA As New ClsLineArt
Private SScrollPic As New ClsFramedScrollPicture
Private DScrollPic As New ClsFramedScrollPicture

Private Sub BtnGreyScale_Click()

    ButtonState False
    LA.GreyScaleImage
    ButtonState True

End Sub

Private Sub ButtonState(State As Boolean)

  ' turn buttons on/off as needed Demo specific

    Save.Enabled = State
    Load.Enabled = State
    'There has to be a picture loaded
    StartLineArt.Enabled = State And OSF.Loaded
    ' at least one pixel must be active for EdgeDetection
    StartEdgeDetect.Enabled = State And OSF.Loaded And LA.PixelSet > 0
    BtnGreyScale.Enabled = State And OSF.Loaded
    Paste2Source.Enabled = State And DScrollPic.ContainsImage
    Save.Enabled = Paste2Source.Enabled
    CancelDraw.Enabled = Not State

    If Paste2Source.Enabled = False And State And OSF.Loaded Then
        ContiansPictureOverRide.Enabled = Not DScrollPic.ContainsImage
      Else 'NOT PASTE2SOURCE.ENABLED...
        ContiansPictureOverRide.Enabled = False
    End If

End Sub

Private Sub CancelDraw_Click()

    LA.Cancel = True ' Cancel The Draw

End Sub

Private Sub Check1_Click()

    LA.ChangeWhileDrawing = Check1.Value = vbChecked

End Sub

Private Sub chkPixel_Click(Index As Integer)

  Dim Val As Integer

    Val = 2 ^ (Index + 1)
    LA.PixelSet = LA.PixelSet + IIf(chkPixel(Index).Value = vbChecked, Val, -Val)
    ButtonState True

End Sub

Private Sub Combo1_Click()

    LA.AnyPixel = Combo1.ListIndex = 0

End Sub

Private Sub Command1_Click()

    HelpForm.Show , Me

End Sub

Private Sub Command2_Click(Index As Integer)

    ButtonState False
    LA.Experimental Index
    ButtonState True

End Sub

Private Sub ContiansPictureOverRide_Click()

  'Override ContiansPicture so that the 'Paste to Source' and 'Save Line Art Image' buttons
  'can be turned on for mostly white images

    DScrollPic.ContainsImage = True
    ButtonState True

End Sub

Private Sub Form_Load()

  'Set Form controls initial values

    Combo1.AddItem "Any"
    Combo1.AddItem "All"
    ''Arvinda's original used the settings
    Combo1.ListIndex = 0
    chkPixel(1).Value = vbChecked
    chkPixel(7).Value = vbChecked
    Tolerance.Value = 157

    'set the various classes (order is not important for the AssignXXXX
    DScrollPic.AssignControls Dest, "Line Art Picture"
    DScrollPic.AssignInterlockScrolls Source
    SScrollPic.AssignControls Source, "Original Picture" ', SScrollV, SScrollH
    SScrollPic.AssignInterlockScrolls Dest
    'the next line can only be done after the Detination SCrollablePictureBox has been set up
    DScrollPic.ResizeTo Source
    With OSF
        .InitDlgs 'initalize save and open dialogs
        .SaveTitle = "Save Line Art Image..."
        .OpenTitle = "Load Colour Image..."
        .UseGraphicFilters = True
        '.FilterIndex = 2' uncomment to default to BMP
    End With 'OSF
    LA.AssignControls Source, Dest, PercentDone
    IgnoreTol.Value = vbChecked 'turn tolerence off for Greyscaling
    greys.Value = 255
    ButtonState True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Do While OSF.Saving  ' wait until the app has finished saving a file.
        DoEvents
    Loop
    Unload Me 'unload
    End       'end

End Sub

Private Sub greys_Change()

    LA.greys = greys.Value
    grys.Caption = LA.greys

End Sub

Private Sub huh_Click()

    LA.GreyWeirdness = huh.Value = 1

End Sub

Private Sub IgnoreTol_Click()

    LA.IgnoreTolerance = IgnoreTol.Value = 1

End Sub

Private Sub Invert_Click()

    LA.Invert = Invert.Value = 1

End Sub

Private Sub KeepOneColorOn(OptB As CheckBox)

  'make sure at least one colour band is being used
  ' this controls the visual display; PreventAllOff in clsLineArt
  ' does this work in the class but gives no feedback to the controls

    If uBlue.Value = 0 And uRed.Value = 0 And uGreen.Value = 0 Then
        OptB.Value = 1
    End If

End Sub

Private Sub Label3_Click()

  'The messy layout in this frame is just to show off the PositionElements and ResizeTo routines
  'in CLsFramedScrollPicture. This routine solves the problem of making the two picture systems
  'identical in size. Set the primary Frame (<-- that one over there) to the size you want, then
  'create the second frame, Setting its Top-Left position where you want it then just put the other
  'controls anywhere on the frame. The routines take care of the rest of your layout.
  'This Label (Label3) is not seen in the program and can be deleted.

End Sub

Private Sub Load_Click()

    OSF.Load_Picture Me, Source
    'reset scroll move values and Dest picture sizes to match Source
    'SScrollpic.NewPicture calls SetMoveValues for SScrollPic
    'DScrollPic.ResizeTo calls SetMoveValues for DScrollPic

    If OSF.Loaded Then
        SScrollPic.NewPicture "Original Image: " & OSF.FilenameOnly
        DScrollPic.ResizeTo Source ', SScrollV, SScrollH
        ButtonState True
        FramOptions.Enabled = True
    End If

End Sub

Private Sub Paste2Source_Click()

    Source.Picture = LoadPicture()
    Set Source.Picture = Dest.Image
    Source.Refresh

End Sub

Private Sub Save_Click()

    OSF.Save_Picture Me, Dest

End Sub

Private Sub StartEdgeDetect_Click()

    ButtonState False
    LA.EdgeDetect
    ButtonState True

End Sub

Private Sub StartLineArt_Click()

    ButtonState False
    LA.LineArt
    ButtonState True

End Sub

Private Sub Text1_Change()

    LA.GreyMode = Val(Text1.Text)

End Sub

Private Sub Tolerance_Change() ' |--Update the Tolerance.

    LA.Tolerance = Tolerance.Value
    Tol.Caption = LA.Tolerance

End Sub

Private Sub Tolerance_Scroll()

    LA.Tolerance = Tolerance.Value
    Tol.Caption = LA.Tolerance

End Sub

Private Sub uBlue_Click()

    KeepOneColorOn uBlue
    LA.Blue = (uBlue.Value = 1)

End Sub

Private Sub uGreen_Click()

    KeepOneColorOn uGreen
    LA.green = (uGreen.Value = 1)

End Sub

Private Sub uRed_Click()

    KeepOneColorOn uRed
    LA.Red = (uRed.Value = 1)

End Sub

':) Ulli's VB Code Formatter V2.13.6 (27/12/2002 8:39:06 PM) 15 + 262 = 277 Lines
