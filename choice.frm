VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save to document"
   ClientHeight    =   3840
   ClientLeft      =   720
   ClientTop       =   1290
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Text            =   "C:\Documents and Settings\Brain\Desktop\not yet listed\"
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Seperate Text with ;"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Text            =   "c:\Documents and Settings\Brain\Desktop\"
      Top             =   2040
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "C:\Documents and Settings\Bri"
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "QUIT"
      DownPicture     =   "choice.frx":0000
      Height          =   855
      Left            =   2880
      Picture         =   "choice.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MAIN"
      DownPicture     =   "choice.frx":0884
      Height          =   855
      Left            =   1440
      Picture         =   "choice.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      DownPicture     =   "choice.frx":1108
      Height          =   855
      Left            =   0
      Picture         =   "choice.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Both"
      DownPicture     =   "choice.frx":198C
      Height          =   1095
      Index           =   2
      Left            =   3000
      Picture         =   "choice.frx":2256
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Microsoft Access"
      DownPicture     =   "choice.frx":2B20
      Height          =   1095
      Index           =   1
      Left            =   1560
      Picture         =   "choice.frx":2F62
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Text Document"
      DownPicture     =   "choice.frx":33A4
      Height          =   1095
      Index           =   0
      Left            =   120
      Picture         =   "choice.frx":37E6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Database location"
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Text location"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Album Number"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Which format do you want to save the information in."
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As Integer
Dim ALBNO
Dim pmno As Integer

Function GetBrowseDirectory(Owner As Form) As String
   Dim bi As BROWSEINFO
   Dim IDL As ITEMIDLIST
   Dim r As Long
   Dim pidl As Long
   Dim tmpPath As String
   Dim pos As Integer

   bi.hOwner = Owner.hwnd
   bi.pidlRoot = 0&
   bi.lpszTitle = "Choose a directory from the list."
   bi.ulFlags = BIF_RETURNONLYFSDIRS
   pidl = SHBrowseForFolder(bi)

   tmpPath = Space$(512)
   r = SHGetPathFromIDList(ByVal pidl, ByVal tmpPath)

   If r Then
      pos = InStr(tmpPath, Chr$(0))
      tmpPath = Left(tmpPath, pos - 1)

      If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
         GetBrowseDirectory = tmpPath
      Else
         GetBrowseDirectory = ""
      End If

End Function



Private Sub Command1_Click()

of = 0
ALBNO = Text1
If ALBNO = "" Then r = MsgBox("Please name the current CD", , "No name"): GoTo ao

If Len(Dir$(Text2 + ALBNO + ".txt")) > 0 And op <> 1 Then r = MsgBox(ALBNO + " already exists. Do you wish to overight?", vbYesNo, "File exists already")
If r = 7 Then GoTo ao
If op = 0 Then dtt
If op = 1 Then sd = 1: frmTable1.Show
If op = 2 Then dtt: sd = 1: frmTable1.Show
Open App.Path & "\config.ini" For Output As 1
Print #1, "Mistalista cofig in V1"
Print #1, "filetxt=" + Text2
Print #1, "filedb=" + Text3
If Mid(Text1, 1, 2) = "PM" Then Print #1, "previpm=" + Str$(Val(Mid$(Text1, 3)) + 1) Else Print #1, "previpm=" + Val(Mid$(Text1, 3))
Close 1
Form1.Enabled = True
Form1.Show
Form4.Hide
ao:
End Sub
Private Sub dtt()
For wot = 0 To gnod
If comments(wot) = "" Then comments(wot) = "No"
lc = Len(comments(wot))
l1 = Len(artists(wot))
l2 = Len(titles(wot))
If lc > lcb Then lcd = lc
If l1 > l1b Then l1b = l1
If l2 > l2b Then l2b = l2
Next wot
cl = lc
da = l1b
dt = l2b

Open Text2 + ALBNO + ".txt" For Output As #1

For cownt = 0 To gnod
Rem If fault(cownt) = 1 Then ok = "OK" Else ok = "N/A"
Rem If fault(cownt) = 2 Then falt = "Fault" Else falt = "N/A"
If fault(cownt) = 1 Then ok = "true " Else ok = "false"
If fault(cownt) = 2 Then falt = "true " Else falt = "false"
If Len(comments(cownt)) < cl Then cs = cl - Len(comments(cownt)): For ast = 1 To cs: cts = cts + " ": Next ast

If Len(artists(cownt)) < da Then cs = da - Len(artists(cownt)): For ast = 1 To cs: ats = ats + " ": Next ast
If Len(titles(cownt)) < dt Then cs = dt - Len(titles(cownt)): For ast = 1 To cs: tts = tts + " ": Next ast
Rem If op = 1 Then Print #1, ok; " "; artists(cownt); ats; Tab; comments(cownt); cts; Tab; titles(cownt); tts; Tab; Str(front(cownt)); Tab; Str(back(cownt)); Tab; Str(nocds(cownt)); Tab; ALBNO; Tab; falt
 Rem Print #1, ok; " "; falt; " "; artists(cownt); ats; Tab; titles(cownt); tts; Tab; Str(front(cownt)); Tab; Str(back(cownt)); Tab; Str(nocds(cownt)); Tab; ALBNO; Tab; comments(cownt)
If Check1.Value = 1 Then ass = ";"
 Print #1, artists(cownt); ats; ass; Tab; titles(cownt); tts; ass; Tab; Str(nocds(cownt))

ats = ""
tts = ""
falt = ""
Next cownt

Close #1

End Sub
Private Sub Command2_Click()
Open App.Path & "\config.ini" For Output As 1
Print #1, "Mistalista cofig in V1"
Print #1, "filetxt=" + Text2
Print #1, "filedb=" + Text3
Print #1, "previpm=" + Mid$(Text1, 3)
Close 1
Form1.Enabled = True
Form1.Show
Form4.Hide

End Sub

Private Sub Command3_Click()
Open App.Path & "\config.ini" For Output As 1
Print #1, "Mistalista cofig in V1"
Print #1, "filetxt=" + Text2
Print #1, "filedb=" + Text3
Print #1, "previpm=" + Mid$(Text1, 3)
Close 1
End
End Sub

Private Sub Command4_Click()
  
  Dim myDir As String
   t2 = GetBrowseDirectory(Form4)
If t2 <> "" Then Text2 = t2
  
End Sub

Private Sub Command5_Click()
  Dim myDir As String
   t3 = GetBrowseDirectory(Form4)
If t3 <> "" Then Text3 = t3
  
End Sub

Private Sub Form_Activate()

Open App.Path & "\config.ini" For Input As 1
For I = 0 To 3
Input #1, A
If I = 0 Then If Not A = "Mistalista cofig in V1" Then Close 1: MsgBox ("Invalid config file format"): GoTo skp
If I = 1 Then If StrComp(A, "filetxt=", vbTextCompare) = 1 Then Text2 = Mid$(A, 9)
If I = 2 Then If StrComp(A, "filedb=", vbTextCompare) = 1 Then Text3 = Mid$(A, 8)
If I = 3 Then If StrComp(A, "previpm=", vbTextCompare) = 1 Then Text1 = "PM" & Str(Val(Mid$(A, 9))): Text1 = Replace(Text1, " ", "", 1, , vbTextCompare)
Next I
Close 1
skp:
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Option1_Click(Index As Integer)
op = Index
End Sub

Private Sub Text1_Change()
If InStr(1, Text1, "/", vbTextCompare) Then Text1 = Mid$(Text1, 1, Len(Text1) - 1)
If InStr(1, Text1, "\", vbTextCompare) Then Text1 = Mid$(Text1, 1, Len(Text1) - 1)
If InStr(1, Text1, ":", vbTextCompare) Then Text1 = Mid$(Text1, 1, Len(Text1) - 1)
If InStr(1, Text1, "*", vbTextCompare) Then Text1 = Mid$(Text1, 1, Len(Text1) - 1)
If InStr(1, Text1, "?", vbTextCompare) Then Text1 = Mid$(Text1, 1, Len(Text1) - 1)
If InStr(1, Text1, Chr$(34), vbTextCompare) Then Text1 = Mid$(Text1, 1, Len(Text1) - 1)
If InStr(1, Text1, "<", vbTextCompare) Then Text1 = Mid$(Text1, 1, Len(Text1) - 1)
If InStr(1, Text1, ">", vbTextCompare) Then Text1 = Mid$(Text1, 1, Len(Text1) - 1)
If InStr(1, Text1, Chr$(124), vbTextCompare) Then Text1 = Mid$(Text1, 1, Len(Text1) - 1)
Text1.SelStart = Len(Text1)


End Sub



