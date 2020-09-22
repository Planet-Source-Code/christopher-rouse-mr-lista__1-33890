VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory scan"
   ClientHeight    =   4830
   ClientLeft      =   1440
   ClientTop       =   1380
   ClientWidth     =   11670
   ControlBox      =   0   'False
   Icon            =   "serch dir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785.14
   ScaleMode       =   0  'User
   ScaleWidth      =   11670
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add to Document"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Player >"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3720
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   10800
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "serch dir.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "serch dir.frx":1094
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "serch dir.frx":1CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "serch dir.frx":2938
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   7080
      TabIndex        =   9
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tracks"
      TabPicture(0)   =   "serch dir.frx":358A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "List3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Info"
      TabPicture(1)   =   "serch dir.frx":35A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblIsVBR"
      Tab(1).Control(1)=   "lblModeExtension"
      Tab(1).Control(2)=   "lblChannelMode"
      Tab(1).Control(3)=   "lblPadded"
      Tab(1).Control(4)=   "lblMPEGVersion"
      Tab(1).Control(5)=   "lblFrequency"
      Tab(1).Control(6)=   "lblOriginal"
      Tab(1).Control(7)=   "lblProtected"
      Tab(1).Control(8)=   "lblLayer"
      Tab(1).Control(9)=   "lblBitrate"
      Tab(1).Control(10)=   "lbl1(10)"
      Tab(1).Control(11)=   "lbl1(9)"
      Tab(1).Control(12)=   "lbl1(8)"
      Tab(1).Control(13)=   "lbl1(7)"
      Tab(1).Control(14)=   "lbl1(6)"
      Tab(1).Control(15)=   "lbl1(5)"
      Tab(1).Control(16)=   "lbl1(4)"
      Tab(1).Control(17)=   "lbl1(3)"
      Tab(1).Control(18)=   "lbl1(2)"
      Tab(1).Control(19)=   "lbl1(1)"
      Tab(1).Control(20)=   "Label6"
      Tab(1).Control(21)=   "Label7"
      Tab(1).Control(22)=   "Label5"
      Tab(1).Control(23)=   "Label4"
      Tab(1).Control(24)=   "Label3"
      Tab(1).Control(25)=   "Label1"
      Tab(1).Control(26)=   "Label2"
      Tab(1).ControlCount=   27
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000007&
         Height          =   3346
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblIsVBR 
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   -74280
         TabIndex        =   37
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblModeExtension 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   -73560
         TabIndex        =   36
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lblChannelMode 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   -72600
         TabIndex        =   35
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblPadded 
         BackStyle       =   0  'Transparent
         Caption         =   "Padded"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   -74160
         TabIndex        =   34
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblMPEGVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "MPEG Version"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   -74160
         TabIndex        =   33
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblFrequency 
         BackStyle       =   0  'Transparent
         Caption         =   "Frequency"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   -74040
         TabIndex        =   32
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblOriginal 
         BackStyle       =   0  'Transparent
         Caption         =   "Original"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   -74280
         TabIndex        =   31
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblProtected 
         BackStyle       =   0  'Transparent
         Caption         =   "Protected"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   -74040
         TabIndex        =   30
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblLayer 
         BackStyle       =   0  'Transparent
         Caption         =   "Layer"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   -72600
         TabIndex        =   29
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblBitrate 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitrate"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   -74160
         TabIndex        =   28
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Is VBR :"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   10
         Left            =   -74880
         TabIndex        =   27
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mode Extension :"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   9
         Left            =   -74880
         TabIndex        =   26
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mode:"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   8
         Left            =   -73080
         TabIndex        =   25
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Padding :"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   7
         Left            =   -74880
         TabIndex        =   24
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   6
         Left            =   -74880
         TabIndex        =   23
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Frequency :"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   5
         Left            =   -74880
         TabIndex        =   22
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Original :"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   4
         Left            =   -74880
         TabIndex        =   21
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Protected :"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   3
         Left            =   -74880
         TabIndex        =   20
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Layer :"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   2
         Left            =   -73080
         TabIndex        =   19
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bit Rate :"
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   1
         Left            =   -74880
         TabIndex        =   18
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "comment"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "gener"
         Height          =   255
         Left            =   -73800
         TabIndex        =   16
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "year"
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Album"
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "artist"
         Height          =   375
         Left            =   -74880
         TabIndex        =   13
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Time"
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11040
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "serch dir.frx":35C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "serch dir.frx":5528
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "serch dir.frx":748E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "serch dir.frx":93F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "serch dir.frx":B35A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "serch dir.frx":D2C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   11040
      Top             =   1800
   End
   Begin VB.CommandButton Command9 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1429
      ButtonWidth     =   1773
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "A"
            Object.ToolTipText     =   "Back"
            Object.Tag             =   "a"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Parent"
            Key             =   "B"
            Object.ToolTipText     =   "Parent"
            Object.Tag             =   "b"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Send to tray"
            Key             =   "c"
            Object.ToolTipText     =   "Send to tray"
            Object.Tag             =   "c"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin MediaPlayerCtl.MediaPlayer player 
         Height          =   255
         Left            =   10200
         TabIndex        =   8
         Top             =   360
         Width           =   255
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   -1  'True
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
   End
   Begin VB.CommandButton command2 
      Caption         =   "Scan"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove list"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   6975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   4920
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   38
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dir`s"
         Object.Width           =   7937
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sDirectory As String
Dim sFile As String
Dim dDirectory As String
Dim dFile As String
Dim mDirectory As String
Dim mFile As String
Dim tsDirectory As String
Dim tsFile As String
Dim I As Integer
Dim driv As String
Dim title() As String
Dim Artist() As String
Dim nocd() As Integer
Dim A As Integer
Dim TsDirList() As String
Dim sDirList() As String
Dim mDirList() As String
Dim nod As Integer
Dim fr As Integer
Dim playp
Dim playc
Dim FileName As String
Dim HasTag As Boolean
Dim Tagg As String * 3
Dim tSongname As String * 30
Dim tArtist As String * 30
Dim tAlbum As String * 30
Dim tYear As String * 4
Dim tComment As String * 30
Dim genreTag As String * 1
Dim tgenre
Dim aii
Dim MP3Info As clsMP3Info
Dim filenum As Integer
Dim strBuffer As String






Private Sub Command1_Click()
res = MsgBox("Are you sure you want to quit?", vbOKCancel + vbCritical, "Are You Sure?")
Text1 = res
If res = 1 Then endSplash.Show: Form1.Hide
End Sub




Private Sub Command2_Click()
Form2.Command11.Enabled = True
Form2.Command12.Enabled = True

Command5.Enabled = True
For cd = 0 To 50
front(cd) = 0: back(cd) = 0: comments(cd) = "": fault(cd) = 0
For ci = 0 To 15
picg(cd, ci) = ""
Next ci
Next cd
skp:
For fs = 0 To nod
If ListView1.ListItems.Item(fs + 1).Checked = True Then rmit (fs): GoTo skp
Next fs

For f = 0 To nod
fs = f
Command5.Enabled = True
ReDim Preserve Artist(f), title(f), nocd(f)
Artist(f) = "": title(f) = "": nocd(f) = 0
Call splitfn(sDirList(f), fs)
Call countcd(sDirList(f), fs)

Next f
List2.Clear
Command4.Visible = True
For f = 0 To nod
List2.AddItem Str$(f) + " arist: " + (Artist(f) + " Title: " + title(f) + " NOS: " + Str$(nocd(f)))
Next f
If Err = 9 Then GoTo skip
Form1.Height = 8010
skip:
End Sub
Private Sub rmit(ar)
Dim Tsirist()
ListView1.ListItems.Remove (ar + 1)
nv = 0
For rs = 0 To nod
If rs = ar Then GoTo skp
ReDim Preserve Tsirist(nv)
Tsirist(nv) = sDirList(rs)
nv = nv + 1
skp:
Next rs
For rs2 = 0 To nv - 1
ReDim Preserve sDirList(rs2)
sDirList(rs2) = Tsirist(rs2)
Next rs2
nod = nv - 1: gnod = nv - 1
End Sub
Private Sub countcd(fn, ta)
tsDirectory = fn
tsFile = Dir$(tsDirectory & "/*. ", 16)
pi = 1
        ti = 0
tnod = 0
        Do While Len(tsFile)
            ReDim Preserve TsDirList(I)
 
 If UCase(GetExtension(tsFile)) = ".JPG" Then picg(ta, pi) = tsDirectory & "/" & tsFile: pi = pi + 1
  If UCase(GetExtension(tsFile)) = "JPEG" Then picg(ta, pi) = tsDirectory & "/" & tsFile: pi = pi + 1
   If UCase(GetExtension(tsFile)) = ".BMP" Then picg(ta, pi) = tsDirectory & "/" & tsFile: pi = pi + 1
        If UCase(GetExtension(tsFile)) = ".GIF" Then picg(ta, pi) = tsDirectory & "/" & tsFile: pi = pi + 1
            If GetAttr(dDirectory & "/" + dFile) = 16 Then nc = nc + 1

If GetAttr(tsDirectory & "/" & tsFile) = 16 Then TsDirList(I) = tsDirectory & "/" & tsFile: ti = ti + 1
If GetAttr(tsDirectory & "/" & tsFile) = 17 Then TsDirList(I) = tsDirectory & "/" & tsFile: ti = ti + 1

           
skp:

tsFile = Dir$
        Loop
tnod = ti - 1
If tnod > 1 Then tnod = tnod - 1
If tnod > 4 Then MsgBox "Warning!  This Directory contains more than 4 Directorys please check cd manually", vbCritical, "OOPS! " + tsDirectory & "/" & tsFile
ReDim Preserve nocd(ta): nocd(ta) = tnod
ti = 0
End Sub

Private Sub splitfn(fn, ta)
On Error Resume Next
l = Len(fn) - 2
 For p = l To 1 Step -1
If Mid$(fn, p, 1) = "/" Then ca = p: GoTo cs
Next p
cs:
If ca > 1 Then ma = l - ca + 5: fnr = Right$(fn, ma)

fnc = ""
fnc = Mid(fnr, 4, Len(fnr) - 2)
If l - 3 > 7 Then If Mid$(fnc, Len(fnc) - 5, 6) = "_2CD's" Or Mid$(fnc, Len(fnc) - 5, 6) = "_3CD's" Or Mid$(fnc, Len(fnc) - 5, 6) = "_2CD's" Then fnc = Mid$(fnc, 1, Len(fnc) - 6): aew = "I think this album could be a Various?"

If (InStr(1, Mid$(fnc, 1, 5), "-", vbTextCompare)) > 0 Then fnc = Mid$(fnc, 6, Len(fnc) - 3)
st:
l = Len(fnc)

lo = InStr(1, fnc, "_", vbTextCompare)
If lo > 0 Then GoTo us
lo1 = InStr(1, fnc, "-", vbTextCompare)
If lo1 > 0 Then GoTo ud
txt = InputBox("Sorry I Just Dont Understand this file name." + Chr$(13) + "Please modifiy the below Box, seperate the name and artist with a " + Chr$(34) + "_" + Chr$(34) + " here is an Example:-" + Chr(13) + "enya_The Memory of Trees" + Chr(13) + Chr(13) + aew, "Stupit Programing.", fnc)
If txt <> "" Then fnc = txt: GoTo st Else ReDim Preserve Artist(ta): Artist(ta) = fnc: GoTo scnd
GoTo scnd
Rem underscore
us:
If lo = 0 Then GoTo scnd
SSC = 1
For sc = lo + 1 To l
los = Mid$(fnc, sc, 1)
If los = "_" Then idu = idu + 1
If idu = 1 And SSC = 1 Then lout = sc: SSC = 0

Next sc
idu = idu + 1
If idu = 1 Then art = Mid(fnc, 1, lo - 1): tit = Mid$(fnc, lo + 1, l - lo): ReDim Preserve Artist(ta), title(ta): Artist(ta) = art: title(ta) = tit:
If lout > 1 Then fnc = Replace(fnc, "_", " ", 1, idu, vbTextCompare)
If lout > 1 Then art = Mid(fnc, 1, lout - 1): tit = Mid$(fnc, lout + 1, l - lout): ReDim Preserve Artist(ta), title(ta): Artist(ta) = art: title(ta) = tit


Rem end underscor
Rem dash
ud:
If lo1 = 0 Then GoTo scnd
SSC = 1
For sc = lo1 + 1 To l
losd = Mid$(fnc, sc, 1)
If losd = "-" Then idD = idD + 1
If idD = 1 And SSC = 1 Then lout2 = sc: SSC = 0

Next sc
idD = idD + 1
 
If idD = 1 Then art = Mid(fnc, 1, lo1 - 1): tit = Mid$(fnc, lo1 + 1, l - lo1): ReDim Preserve Artist(ta), title(ta): Artist(ta) = art: title(ta) = tit
If lout2 > 1 Then fnc = Replace(fnc, "-", " ", 1, idD, vbTextCompare)

If lout2 > 1 Then art = Mid(fnc, 1, lout2 - 1): tit = Mid$(fnc, lout2 + 1, l - lout2): ReDim Preserve Artist(ta), title(ta): Artist(ta) = art: title(ta) = tit

Rem end dash

scnd:
Text3 = fnc
End Sub





Private Sub command2_GotFocus()
command2.FontBold = True
End Sub

Private Sub command2_LostFocus()
command2.FontBold = False
End Sub

Private Sub command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
command2.SetFocus
End Sub

Private Sub Command3_Click()
Call Form_Load
End Sub



Private Sub Command4_Click()
Command4.Visible = False
Form1.Height = 5220
End Sub



Private Sub Command5_Click()
For cas = 0 To nod
titles(cas) = title(cas)
artists(cas) = Artist(cas)
nocds(cas) = nocd(cas)

Next cas
Form4.Show
Form1.Hide
Form1.Enabled = False

End Sub



Private Sub Command6_Click()
If Form1.Width = 7095 Then Form1.Width = 10815: Command6.Caption = "Player <": Call ListView1_Click Else Form1.Width = 7095: Command6.Caption = "Player >"
End Sub


Private Sub Command7_Click()
player.Stop
playc = 1
End Sub



Private Sub Command8_Click()
playc = 0
If List3.ListIndex = List3.ListCount Then GoTo avoid
List3.ListIndex = List3.ListIndex - 1
player.Open mDirList(List3.ListIndex)
avoid:
End Sub



Private Sub Command9_Click()
playc = 0
If List3.ListIndex = List3.ListCount - 1 Then GoTo avoid
List3.ListIndex = List3.ListIndex + 1
player.Open mDirList(List3.ListIndex)
GetTag (mDirList(List3.ListIndex))

avoid:
End Sub



Private Sub Drive1_Change()
Command5.Enabled = True
On Error GoTo er
rt:


driv = Drive1.Drive
driv = Mid(driv, 1, 2)
I = 0


Call Form_Load

er:
If Err = 52 Then e = Err: res = MsgBox("Please insert a CD\disk in drive " + UCase(driv), vbRetryCancel + vbCritical, "Disk Error")
If res = 4 Then Call Drive1_Change
If res = 2 Then Drive1.Drive = "C:": Drive1_Change
End Sub







Private Sub Form_Activate()
Label3.Caption = "Artist: " + tArtist
Label4.Caption = "Album: " + tAlbum
Label1.Caption = "Title: " + tSongname
Label5.Caption = "Year: " + tYear
Label6.Caption = "Comment: " + tComment
Label7.Caption = "Genre: " + tgenre
If sf = 1 And cwo = 1 Then A = ind
If sf <> 1 Then GoTo A
If Artist(A) <> artg Then Artist(A) = artg
If title(A) <> titg Then title(A) = titg
If nocd(A) <> nocdg Then nocd(A) = Str$(nocdg)

List2.Clear
For f = 0 To nod
List2.AddItem Str$(f) + " arist: " + (Artist(f) + " Title: " + title(f) + Str$(nocd(f)))
Next f

If cwo <> 1 Then GoTo A

If sf = 1 And cwo = 1 And iod = "+" Then ind = ind + 1
If sf = 1 And cwo = 1 And iod = "-" Then ind = ind - 1

A = ind
If A < 1 Then Form2.Command11.Enabled = False
If A < 0 Then A = 0: Form2.Command11.Enabled = False
If A > nod - 1 Then Form2.Command12.Enabled = False
If A > nod Then A = nod: Form2.Command12.Enabled = False
pnum = ind
iod = ""
cwo = 0
pnum = Val(A)

artg = Artist(A)
titg = title(A)
nocdg = nocd(A)

Form2.Enabled = True
Form2.Show
Form1.Enabled = False

A:

sf = 0
End Sub

Private Sub Form_Load()
If aii = 1 Then GoTo aii
SubClass Me.hwnd
    
    '------------------------------------------
    'Preset portions of the data structure
    '------------------------------------------
    structNotify.lStructureSize = 88&
    structNotify.lFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    structNotify.lCallBackMessage = UM_TASKBARMESSAGE
    structNotify.sTip = ""
    structNotify.hwnd = Me.hwnd
Set MP3Info = New clsMP3Info
aii = 1
aii:
Command5.Enabled = False
If fr = 0 Then Form1.Width = 7095: fr = 1
If driv = "" Then driv = "c:"
sDirectory = driv

ListView1.ListItems.Clear

sFile = Dir$(sDirectory & "/*. ", vbDirectory)

        I = 0

        Do While Len(sFile)

            ReDim Preserve sDirList(I)
GAT = 0
If Trim$(sFile) = ".." Then GoTo skp
If Trim$(sFile) = "." Then GoTo skp

If GetAttr(sDirectory & "/" + sFile) = 16 Then GAT = 1
If GetAttr(sDirectory & "/" + sFile) = 17 Then GAT = 1
If GAT = 0 Then GoTo skp
sDirList(I) = sDirectory & "/" & sFile

            I = I + 1
skp:
sFile = Dir$
       
        Loop
nod = I - 1
gnod = nod
I = 0
For I = 0 To gnod
dDirectory = sDirList(I)



dFile = Dir$(dDirectory & "/*. ", vbDirectory)
nc = 0
     
        Do While Len(dFile)

If GetAttr(dDirectory & "/" + dFile) = 16 Then nc = nc + 1
If GetAttr(dDirectory & "/" + dFile) = 17 Then nc = nc + 1



apl:
dFile = Dir$
       
        Loop

For cds = 1 To Len(sDirList(I)) Step 1
If Mid$(sDirList(I), cds, 1) = "/" Then cmdl = Mid$(sDirList(I), cds + 1)
Next cds
nc = nc - 1
If nc < 1 Then nc = 1
If nc > 6 Then nc = 6

lis = ListView1.ListItems.Add(1 + I, , cmdl, nc, nc)
Rem MsgBox sDirList(I) + " " + Str$(nc)
sf:
nc = 0

Next I
ListView1.Refresh

End Sub




Public Function GetExtension(lpFileName As String)
'parses extension from filename and returns the extension
'no extension is sent back as """"

    Dim nPosition As Integer
    
    If InStr(1, lpFileName, ".") < 1 Then
     Exit Function
     Else
        nPosition = InStr(1, lpFileName, ".")
      lt = Len(lpFileName)
      If lt - nPosition > 4 Then GetExtension = Mid(lpFileName, lt - 3, 4): Exit Function
      
      
        
        GetExtension = Mid$(lpFileName, _
            nPosition, _
            Len(lpFileName) - (nPosition - 1))
    End If
End Function
Private Sub List1_Click()
playc = 1
Rem If player.PlayState <> 0 Then player.Stop
List3.Clear
If Form1.Width < 10815 Then GoTo skp
mDirectory = Mid$(driv, 1, 2) + List1 + "/"
mDirectory = driv + "/" + List1 + "/"
mFile = Dir$(mDirectory & "*.mp3", 16)

        I = 0

        Do While Len(mFile)
            ReDim Preserve mDirList(I)
       mDirList(I) = mDirectory & "" & mFile
        std = 0
        For cds = Len(mDirList(I)) To 1 Step -1
        If Mid$(mDirList(I), cds, 1) = "/" And std = 0 Then smdl = Right$(mDirList(I), Len(mDirList(I)) - cds): std = 1
        Next cds
        List3.AddItem smdl, I
        I = I + 1
        mFile = Dir$
        Loop
skp:
End Sub

Private Sub List1_dblClick()
On Error Resume Next
driv = sDirList(List1.ListIndex)
Rem l = Len(driv)
Rem ll = Len(listi)
Rem lo = ll - l
Rem List = Right$(listi, lo + 2)
Rem driv = driv + "/" + List
Form_Load
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For bl = 1 To ListView1.ListItems.Count
          If ListView1.ListItems.Item(bl).ForeColor = vbBlue Then ListView1.ListItems.Item(bl).ForeColor = vbBlack: ListView1.ListItems.Item(bl).Bold = False: ListView1.ListItems.Item(bl).Ghosted = False
        
       Next bl
End Sub

Private Sub List2_Click()
A = (List2.ListIndex)
If A > 0 Then Form2.Command11.Enabled = True
If A < nod Then Form2.Command12.Enabled = True

If A < 1 Then Form2.Command11.Enabled = False
If A > nod - 1 Then Form2.Command12.Enabled = False
pnum = Val(A)
ind = A
artg = Artist(A)
titg = title(A)

nocdg = nocd(A)
Form2.Enabled = True
Form2.Show
Form1.Enabled = False
End Sub

Private Sub List3_Click()
playc = 0
If List3.ListIndex = List3.ListCount Then GoTo avoid
If List3.ListIndex < 0 Then List3.ListIndex = 0
player.Open mDirList(List3.ListIndex)
GetTag (mDirList(List3.ListIndex))
Label3.Caption = "Artist: " + tArtist
Label4.Caption = "Album: " + tAlbum
Label1.Caption = "Title: " + tSongname
Label5.Caption = "Year: " + tYear
Label6.Caption = "Comment: " + tComment
Label7.Caption = "Genre: " + tgenre
DisplayMP3Info (mDirList(List3.ListIndex))
tcv = Val(Mid$(lblBitrate.Caption, 1, 3))
If tcv < 128 Then MsgBox ("This Mp3 is LOWER quality than CD Quality unless this is MP3 Layer II(MP3 Pro)")
avoid:
End Sub



Private Sub ListView1_Click()
Text1 = driv + "/" + ListView1.SelectedItem + "/"
playc = 1
Rem If player.PlayState <> 0 Then player.Stop
List3.Clear
If Form1.Width < 10815 Then GoTo skp
mDirectory = Mid$(driv, 1, 2) + ListView1.SelectedItem + "/"
mDirectory = driv + "/" + ListView1.SelectedItem + "/"
mFile = Dir$(mDirectory & "*.mp3", 16)
        I = 0

        Do While Len(mFile)
            ReDim Preserve mDirList(I)
       mDirList(I) = mDirectory & "" & mFile
      std = 0
        For cds = Len(mDirList(I)) To 1 Step -1
        If Mid$(mDirList(I), cds, 1) = "/" And std = 0 Then smdl = Right$(mDirList(I), Len(mDirList(I)) - cds): std = 1
        Next cds
        List3.AddItem smdl, I
        I = I + 1
        mFile = Dir$
        Loop
skp:
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
ic = ListView1.SelectedItem.Icon
If ic - 1 = 0 Then GoTo skp
driv = sDirList(ListView1.SelectedItem.Index - 1)
Form_Load
skp:
End Sub

Private Sub ListView1_LostFocus()
For bl = 1 To ListView1.ListItems.Count
          If ListView1.ListItems.Item(bl).ForeColor = vbBlue Then ListView1.ListItems.Item(bl).ForeColor = vbBlack: ListView1.ListItems.Item(bl).Bold = False: ListView1.ListItems.Item(bl).Ghosted = False
        
       Next bl
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


    For Il = 1 To ListView1.ListItems.Count 'Goes through all items In the listView
        'checks to see if the mouse is over the
        '     current listView item
        If (X > ListView1.ListItems.Item(Il).Left) And _
        (X < (ListView1.ListItems.Item(Il).Left + ListView1.ListItems.Item(Il).Width)) _
        And (Y > ListView1.ListItems.Item(Il).Top) And _
        (Y < ListView1.ListItems.Item(Il).Top + ListView1.ListItems.Item(Il).Height) Then
        'if it is, set all to default, in this c
        '     ase, black


        For bl = 1 To ListView1.ListItems.Count
          If ListView1.ListItems.Item(bl).ForeColor = vbBlue Then ListView1.ListItems.Item(bl).ForeColor = vbBlack: ListView1.ListItems.Item(bl).Bold = False: ListView1.ListItems.Item(bl).Ghosted = False
        
       Next bl
        'sets the one that the mouse is over to
        '     Blue, can be changed.
        ListView1.ListItems.Item(Il).ForeColor = vbBlue: ListView1.ListItems.Item(Il).Bold = True: ListView1.ListItems.Item(Il).Ghosted = True
    End If
Next Il
End Sub




Private Sub Timer1_Timer()

Dim tinseconden As Long
Dim lengths, lengths1, min, sec As Long
lengths = player.Duration
tinseconden = player.CurrentPosition
lengths1 = lengths - tinseconden
min = lengths1 \ 60
sec = lengths1 - min * 60
Label2.Caption = "Time remaining: " & Val(min) & " : " & Val(sec)
If List3.ListIndex = List3.ListCount - 1 Then GoTo avoid
If player.PlayState = 0 And playc = 0 Then Call Command9_Click
avoid:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "A"
l = Len(driv)
For p = l To 1 Step -1
If Mid$(driv, p, 1) = "/" Then ca = p: GoTo cs
Next p
cs:
If ca > 1 Then driv = Mid$(driv, 1, ca - 1)

Form_Load
Case "B"
driv = Mid$(driv, 1, 2)
Form_Load
Case "c"
AddIcon 1, "MR Lista", Form1.Icon
Form1.Enabled = False: Form1.Hide

End Select
End Sub


Private Sub GetTag(FileName)
Open FileName For Binary As #1
Get #1, FileLen(FileName) - 127, Tagg
If Not Tagg = "TAG" Then
Close #1
HasTag = False
tSongname = "No Tag Found"
tArtist = "No Tag Found"
tAlbum = "No Tag Found"
tYear = "None"
tComment = "No Tag Found"
tgenre = "0"
Exit Sub
End If
HasTag = True
Get #1, , tSongname
Get #1, , tArtist
Get #1, , tAlbum
Get #1, , tYear
Get #1, , tComment
Get #1, , genreTag
If Asc(tArtist) = 0 Then tArtist = "No tag found"
If Asc(tSongname) = 0 Then tSongname = "No tag found"
If Asc(tAlbum) = 0 Then tAlbum = "No tag found"
If Asc(tYear) = 0 Then tYear = "none"
If Asc(tComment) = 0 Then tComment = "No tag found"
If Asc(genreTag) = 0 Then genreTag = "No tag found"

Rem tSongname
Rem tartist
Rem tAlbum
Rem tYear
Rem tcomment
Rem tGenre
Close #1

    genreList(0).genre = "Blues"
    genreList(1).genre = "Classic Rock"
    genreList(10).genre = "New Age"
    genreList(100).genre = "Humour"
    genreList(101).genre = "Speech"
    genreList(102).genre = "Chanson"
    genreList(103).genre = "Opera"
    genreList(104).genre = "Chamber Music"
    genreList(105).genre = "Sonata"
    genreList(106).genre = "Symphony"
    genreList(107).genre = "Booty Brass"
    genreList(108).genre = "Primus"
    genreList(109).genre = "Porn Groove"
    genreList(11).genre = "Oldies"
    genreList(110).genre = "Satire"
    genreList(111).genre = "Slow Jam"
    genreList(112).genre = "Club"
    genreList(113).genre = "Tango"
    genreList(114).genre = "Samba"
    genreList(115).genre = "Folklore"
    genreList(116).genre = "Ballad"
    genreList(117).genre = "Poweer Ballad"
    genreList(118).genre = "Rhytmic Soul"
    genreList(119).genre = "Freestyle"
    genreList(12).genre = "Other"
    genreList(120).genre = "Duet"
    genreList(121).genre = "Punk Rock"
    genreList(122).genre = "Drum Solo"
    genreList(123).genre = "A Capela"
    genreList(124).genre = "Euro-House"
    genreList(125).genre = "Dance Hall"
    genreList(13).genre = "Pop"
    genreList(14).genre = "R and B"
    genreList(15).genre = "Rap"
    genreList(16).genre = "Reggae"
    genreList(17).genre = "Rock"
    genreList(18).genre = "Techno"
    genreList(19).genre = "Industrial"
    genreList(2).genre = "Country"
    genreList(20).genre = "Alternative"
    genreList(21).genre = "Ska"
    genreList(22).genre = "Death Metal"
    genreList(23).genre = "Pranks"
    genreList(24).genre = "Soundtrack"
    genreList(25).genre = "Euro-Techno"
    genreList(26).genre = "Ambient"
    genreList(27).genre = "Trip-Hop"
    genreList(28).genre = "Vocal"
    genreList(29).genre = "Jazz+Funk"
    genreList(3).genre = "Dance"
    genreList(30).genre = "Fusion"
    genreList(31).genre = "Trance"
    genreList(32).genre = "Classical"
    genreList(33).genre = "Instrumental"
    genreList(34).genre = "Acid"
    genreList(35).genre = "House"
    genreList(36).genre = "Game"
    genreList(37).genre = "Sound Clip"
    genreList(38).genre = "Gospel"
    genreList(39).genre = "Noise"
    genreList(4).genre = "Disco"
    genreList(40).genre = "AlternRock"
    genreList(41).genre = "Bass"
    genreList(42).genre = "Soul"
    genreList(43).genre = "Punk"
    genreList(44).genre = "Space"
    genreList(45).genre = "Meditative"
    genreList(46).genre = "Instrumental Pop"
    genreList(47).genre = "InstrumentalRock"
    genreList(48).genre = "Ethnic"
    genreList(49).genre = "Gothic"
    genreList(5).genre = "Funk"
    genreList(50).genre = "Darkwave"
    genreList(51).genre = "Techno-Industrial"
    genreList(52).genre = "Electronic"
    genreList(53).genre = "Pop-Folk"
    genreList(54).genre = "Eurodance"
    genreList(55).genre = "Dream"
    genreList(56).genre = "Southern Rock"
    genreList(57).genre = "Comedy"
    genreList(58).genre = "Cult"
    genreList(59).genre = "Gangsta"
    genreList(6).genre = "Grunge"
    genreList(60).genre = "Top 40"
    genreList(61).genre = "Christian Rap"
    genreList(62).genre = "Pop/Funk"
    genreList(63).genre = "Jungle"
    genreList(64).genre = "Native American"
    genreList(65).genre = "Cabaret"
    genreList(66).genre = "New Wave"
    genreList(67).genre = "Psychadelic"
    genreList(68).genre = "Rave"
    genreList(69).genre = "Showtunes"
    genreList(7).genre = "Hip-Hop"
    genreList(70).genre = "Trailer"
    genreList(71).genre = "Lo-Fi"
    genreList(72).genre = "Tribal"
    genreList(73).genre = "Acid Punk"
    genreList(74).genre = "Acid Jazz"
    genreList(75).genre = "Polka"
    genreList(76).genre = "Retro"
    genreList(77).genre = "Musical"
    genreList(78).genre = "Rock&Roll"
    genreList(79).genre = "Hard Rock"
    genreList(8).genre = "Jazz"
    genreList(80).genre = "Folk"
    genreList(81).genre = "Folk-Rock"
    genreList(82).genre = "National Folk"
    genreList(83).genre = "Swing"
    genreList(84).genre = "Fast Fusion"
    genreList(85).genre = "Bebob"
    genreList(86).genre = "Latin"
    genreList(87).genre = "Revival"
    genreList(88).genre = "Celtic"
    genreList(89).genre = "Bluegrass"
    genreList(9).genre = "Metal"
    genreList(90).genre = "Avantgarde"
    genreList(91).genre = "Gothic Rock"
    genreList(92).genre = "Progressive Rock"
    genreList(93).genre = "Psychedelic Rock"
    genreList(94).genre = "Symphonic Rock"
    genreList(95).genre = "Slow Rock"
    genreList(96).genre = "Big Band"
    genreList(97).genre = "Chorus"
    genreList(98).genre = "Easy Listening"
    genreList(99).genre = "Acoustic"
If genreTag <> "" Then ' if genreTag is a nullstring we skip the Asc function
        intGenreNumber = Asc(genreTag) ' or a run-time error occurs
    Else
        intGenreNumber = 255
    End If
    
    If intGenreNumber > 125 Then ' legal genre tags run from 0 to 125
        tgenre = "Unknown" ' anything else we will call unknown
    Else
        tgenre = genreList(intGenreNumber).genre ' return the string desc of genre
    End If
End Sub
Sub DisplayMP3Info(strMP3Path As String)

Dim NoMoreSlash As Boolean
Dim strTemp As String
Dim nPos As Integer


If Dir(strMP3Path) = "" Then Exit Sub


MP3Info.FileName = strMP3Path
MP3Info.ReadMP3Header

If Not MP3Info.ValidHeader Then
  
    lblMPEGVersion.Caption = MP3Info.ID
    lblLayer.Caption = MP3Info.Layer
    lblBitrate.Caption = MP3Info.Bitrate & " kbps"
    lblChannelMode.Caption = MP3Info.Mode
    lblModeExtension.Caption = MP3Info.ModeExt
    lblOriginal.Caption = MP3Info.Original
    lblPadded.Caption = MP3Info.Padded
    lblProtected.Caption = MP3Info.ProtectionBitSet
    lblFrequency.Caption = MP3Info.Frequency & Hz
    Exit Sub
End If

strTemp = strMP3Path

Do
    nPos = InStr(1, strTemp, "/")
    
    If nPos <> 0 Then
     strTemp = Right$(strTemp, Len(strTemp) - nPos)
    Else
     NoMoreSlash = True
    End If

DoEvents
Loop Until NoMoreSlash = True


'MP3 Headers

lblMPEGVersion.Caption = MP3Info.ID
lblLayer.Caption = MP3Info.Layer
lblBitrate.Caption = MP3Info.Bitrate & " kbps"
lblChannelMode.Caption = MP3Info.Mode
lblModeExtension.Caption = MP3Info.ModeExt
lblOriginal.Caption = MP3Info.Original
lblPadded.Caption = MP3Info.Padded
lblProtected.Caption = MP3Info.ProtectionBitSet
lblFrequency.Caption = MP3Info.Frequency

'MP3 ID3v1.1 info.




End Sub


Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For bl = 1 To ListView1.ListItems.Count
          If ListView1.ListItems.Item(bl).ForeColor = vbBlue Then ListView1.ListItems.Item(bl).ForeColor = vbBlack: ListView1.ListItems.Item(bl).Bold = False: ListView1.ListItems.Item(bl).Ghosted = False
        
       Next bl
End Sub
