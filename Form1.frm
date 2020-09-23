VERSION 5.00
Begin VB.Form HelpForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Line Art Tips"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "HelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Me.Hide

End Sub

Private Sub Form_Load()

    Text1.Text = LineArtTip

End Sub

Private Function LineArtTip() As String

  Dim Str As String

    Str = Str & "This is a simple help file built from Arvinder's original labels with further comments from Roger." & vbNewLine
    Str = Str & "Press [Esc] or click OK to close." & vbNewLine
    Str = Str & "" & vbNewLine
    Str = Str & "Note everything except the 'Help' and 'Load A Colour\Greyscale Image' buttons is disabled when you first start." & vbNewLine
    Str = Str & "" & vbNewLine
    Str = Str & "   Arvinder's Comments and Advice" & vbNewLine
    Str = Str & "If a picture has a lot of one colour in it, then only tick that colour box, and put the tolerance on low." & vbNewLine
    Str = Str & "ie If there is a lot of Red; Tick Red." & vbNewLine
    Str = Str & "Ticking Red, Green and Blue, usually gives a good result." & vbNewLine
    Str = Str & "Experiment with all the controls, and see what you get." & vbNewLine
    Str = Str & "If there is too much Black in the resulting image, then decrease the Tolerance." & vbNewLine
    Str = Str & "E-Mail me at: Arvinder@Sehmi.org.uk if you need help, or have a question." & vbNewLine
    Str = Str & "" & vbNewLine
    Str = Str & "   Roger's Comments and Advice" & vbNewLine
    Str = Str & "This is a modification of Arvinder's original code." & vbNewLine
    Str = Str & "I put everything in classes (ClsOpenSave is also based on a bas module in Arvinda's code)" & vbNewLine
    Str = Str & "and added Min Thant Sin's (contact through PSC) picture scrolling code (also converted to a class)." & vbNewLine
    Str = Str & "I also optimised Arvinder's code for speed and added the Edge Detection options." & vbNewLine
    Str = Str & "I have also added a Grey Scale option." & vbNewLine
    Str = Str & "I changed the layout of the form to allow for the Edge Detection options and increased the size of the picture boxes." & vbNewLine
    Str = Str & "" & vbNewLine
    Str = Str & "   NEW" & vbNewLine
    Str = Str & "The Edge Detection options allow you to change where and how the Edge Detection routine works." & vbNewLine
    Str = Str & "Select the Pixel(s) you want tested on the square of checkboxes" & vbNewLine
    Str = Str & "Select Any (any checked pixel detected is enough) or All (every checked pixel must be detected)" & vbNewLine
    Str = Str & "NOTE 1: At least one Pixel must be set for Edge Detection to work." & vbNewLine
    Str = Str & "NOTE 2: Arvinda's original code equated to Top, Left and Any so those are the defaults set from Form_Load." & vbNewLine
    Str = Str & "NOTE 3:Tolerence is normal locked during Edge Detection (it's faster that way). " & vbNewLine
    Str = Str & "               However for experimental purposes on large images you can " & vbNewLine
    Str = Str & "               Check 'Change Tolerence while drawing' and Tolerence can be changed at will." & vbNewLine
    Str = Str & "NOTE 4: Tolerence works as Arvinda said for Line Art but for Edge Detection lower values give more Black." & vbNewLine
    Str = Str & "                (Also depends on how many pixels you test and whether it is Any or All)." & vbNewLine
    Str = Str & "NOTE 5: If you select all pixels and use All very few pixels will be detected." & vbNewLine
    Str = Str & "" & vbNewLine
    Str = Str & "NEW" & vbNewLine
    Str = Str & "The Grey Scaling system allows you to produce grey scaled images with between 2 and 255 levels of grey." & vbNewLine
    Str = Str & "NOTE 1 Tolerence is by default Off. Turned on it reduces all greys above the Tolerence value to white." & vbNewLine
    Str = Str & "NOTE 2 An original may not contain enough colours to produce the full range of greys you have asked for." & vbNewLine
    Str = Str & "NOTE 3 Due to issue in note 3 you may not be able to see any difference between close grey numbers." & vbNewLine
    Str = Str & "NOTE 4 'Greys' = 1 + 'Ignore Tolerence' = True matches Line Art All Colour Values selected + 'Tolerence' = 128" & vbNewLine
    Str = Str & "" & vbNewLine
    Str = Str & "IMAGES" & vbNewLine
    Str = Str & "You can load any picture into the Original picturebox, but both routines work better with less complex images." & vbNewLine
    Str = Str & "Images can only be saved as BMP because I (and Arvinder) use the VB SavePicture routine." & vbNewLine
    Str = Str & "I have included the test images Arvinder supplied with the original upload and the screenshot I used at PCS." & vbNewLine
    Str = Str & "Try loading an image you have processed and reprocessing it for even more interesting effects (screenshot allows you to test this)." & vbNewLine
    Str = Str & "" & vbNewLine
    Str = Str & "Experimental" & vbNewLine
    Str = Str & "The items in this box are experimental. They are not fully intergrated in the class. Some respond to Tolerance and some use the Pixel Array in Edge Detection. As they are still in development I have not documented them fully so check the code and experiment." & vbNewLine
    Str = Str & "" & vbNewLine
    Str = Str & "FURTHER HELP" & vbNewLine
    Str = Str & "The code is heavily commented, so please read it." & vbNewLine
    Str = Str & "" & vbNewLine
    Str = Str & "Hope you like it." & vbNewLine
    Str = Str & "" & vbNewLine
    Str = Str & "Email: rojagilkrist@hotmail.com" & vbNewLine
    LineArtTip = Str & "" & vbNewLine

End Function

':) Ulli's VB Code Formatter V2.13.6 (27/12/2002 8:38:03 PM) 1 + 76 = 77 Lines
