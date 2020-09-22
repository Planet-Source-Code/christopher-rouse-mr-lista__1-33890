VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Title"
   ClientHeight    =   7350
   ClientLeft      =   2715
   ClientTop       =   3840
   ClientWidth     =   13545
   ControlBox      =   0   'False
   Icon            =   "fnreq.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   13545
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command16 
      Caption         =   "Cvfl"
      Height          =   375
      Left            =   5880
      TabIndex        =   40
      ToolTipText     =   "Change artist text to lower case exept the first letter"
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Cvfl"
      Height          =   375
      Left            =   5880
      TabIndex        =   39
      ToolTipText     =   "Change title text to lower case exept the first letter"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "+ Soundtrack to title"
      Height          =   255
      Left            =   4080
      TabIndex        =   38
      Top             =   4920
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Fault"
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   37
      Top             =   6000
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "OK"
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   36
      Top             =   5640
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Not cheked"
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   35
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      Caption         =   "AUTO"
      Height          =   255
      Left            =   11520
      TabIndex        =   33
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4320
      TabIndex        =   32
      Text            =   "Text7"
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Next"
      Height          =   495
      Left            =   2400
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Previous"
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Add **** to Title"
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Set Artist to Various"
      Height          =   255
      Left            =   2160
      TabIndex        =   28
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   6360
      Width           =   5775
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Back cover"
      Height          =   255
      Left            =   11520
      TabIndex        =   25
      Top             =   480
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Front cover"
      Height          =   375
      Left            =   11520
      TabIndex        =   24
      Top             =   120
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   6960
      TabIndex        =   23
      Top             =   120
      Width           =   4455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   5415
   End
   Begin VB.CommandButton Command8 
      Caption         =   "CFL"
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      ToolTipText     =   "Change artist text to lower case exept the first leter of every word"
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "CFL"
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      ToolTipText     =   "Change title text to lower case exept the first leter of every word"
      Top             =   3240
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      DataSource      =   "split"
      Height          =   315
      ItemData        =   "fnreq.frx":0442
      Left            =   120
      List            =   "fnreq.frx":0444
      TabIndex        =   17
      Top             =   2280
      Width           =   5415
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   960
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "lc"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      ToolTipText     =   "Change artist text to lower case"
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "lc"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      ToolTipText     =   "Change title text to lower case "
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UC"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      ToolTipText     =   "Change artist text to upper case"
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UC"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      ToolTipText     =   "Change title text to upper case"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   11640
      TabIndex        =   9
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Swap title with artist"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   4080
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label10 
      Caption         =   $"fnreq.frx":0446
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8280
      TabIndex        =   34
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   7320
      Top             =   7800
      Width           =   2160
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5535
      Left            =   6960
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Label Label9 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   6000
      Width           =   4455
   End
   Begin VB.Label Label8 
      Caption         =   "Pull down list to show some posible artist names:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label Label7 
      Caption         =   "Pull down list to title some posible title names:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "CD`s.    type in box new Qty."
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "This has "
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Change title to:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Change Artist to:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "I think the artist is:"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "I think the title is :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim split() As String
Dim prv
Dim fs
Dim bs






Private Sub Check1_Click()

front(ind) = Check1.Value
End Sub

Private Sub Check2_Click()

back(ind) = Check2.Value
End Sub

Private Sub Combo1_Click()
Text4 = Combo1
End Sub



Private Sub Combo2_Click()
Text3 = Combo2
End Sub

Private Sub Command1_Click()
 t3 = Text3: t4 = Text4
Text3 = t4: Text4 = t3
End Sub

Private Sub Command10_Click()
If InStr(1, Text3, " ****", vbTextCompare) < 1 Then Text3 = Text3 + " ****"
End Sub

Private Sub Command11_Click()


If ind - 1 < gnod Then Command12.Enabled = True


artg = Text4:  Text2 = Text4
titg = Text3:  Text1 = Text3
comments(ind) = Text6
Text6 = ""

nocdg = Text5
Rem Form2.Hide
Form1.Enabled = True
Form1.Show
iod = "-"

cwo = 1
sf = 1
 Form1.Enabled = True
End Sub

Private Sub Command12_Click()

If ind + 1 > 0 Then Command11.Enabled = True

artg = Text4:  Text2 = Text4
titg = Text3:  Text1 = Text3
comments(ind) = Text6
Text6 = ""

nocdg = Text5
Rem Form2.Hide
Form1.Enabled = True
Form1.Show
iod = "+"
cwo = 1
sf = 1
 Form1.Enabled = True
skp:
End Sub



Private Sub Command13_Click()
If fs = 1 Then Check1 = 1: fs = 0
If bs = 1 Then Check2 = 1: bs = 0

End Sub

Private Sub Command14_Click()
If InStr(1, Text3, " Soundtrack", vbTextCompare) < 1 Then Text3 = Text3 + " Soundtrack"

End Sub

Private Sub Command15_Click()
Text3 = LCase(Text3)
fl = UCase(Mid$(Text3, 1, 1))
Text3 = fl + Mid$(Text3, 2, Len(Text3) - 1)
End Sub

Private Sub Command16_Click()
Text4 = LCase(Text4)
fl = UCase(Mid$(Text4, 1, 1))
Text4 = fl + Mid$(Text4, 2, Len(Text4) - 1)
End Sub

Private Sub Command2_Click()
artg = Text4:  Text2 = Text4
titg = Text3:  Text1 = Text3
comments(ind) = Text6
Text6 = ""

nocdg = Text5
Form2.Hide
Form1.Enabled = True
Form1.Show
sf = 1
Form2.Enabled = False
End Sub


Private Sub Command3_Click()
Text3 = UCase(Text3)
End Sub

Private Sub Command4_Click()
Text4 = UCase(Text4)
End Sub

Private Sub Command5_Click()
Text3 = LCase(Text3)
End Sub

Private Sub Command6_Click()
Text4 = LCase(Text4)
End Sub

Private Sub Command7_Click()
Text3 = StrConv(Text3.Text, vbProperCase)
End Sub

Private Sub Command8_Click()
Text4 = StrConv(Text4.Text, vbProperCase)
End Sub



Private Sub Command9_Click()
Text4 = "Various"
End Sub

Private Sub Form_Activate()
Label10.Caption = "If there is no picture here then click on above list, if there is no pics in the list then ther are no covers in JPG, BMP or GIF Format in the current directory."
On Error GoTo er
If prv = 1 Then prv = 0: GoTo skp
fs = 0
bs = 0
Text7 = ind
Image1.Picture = Image2.Picture
Label10.Visible = True
If sf2 = 1 Then sf2 = 0: GoTo skp
fal = fault(ind)
Option1(fal) = True
Text6 = comments(ind)
Check1 = front(ind)
Check2 = back(ind)
Text1 = titg
Text2 = artg
Text3 = titg
Text4 = artg
Text5 = nocdg
comb = artg + " " + titg
comb = SquishSpaces(comb)
comb = SquishSpaces(comb)

sag:
If ic > 0 Then ReDim Preserve split(ic): split(ic) = wa
l = Len(comb)
csb = 1
For cs = 1 To l
sc = Mid$(comb, cs, 1)

If sc = " " Then ic = ic + 1: wa = Mid$(comb, 1, cs): comb = Right$(comb, l - cs): GoTo sag
Next cs
ic = ic + 1
wa = comb
If ic > 0 Then ReDim Preserve split(ic): split(ic) = wa: ic = ic + 1: ReDim Preserve split(ic): split(ic) = ""
Combo1.Clear
Combo2.Clear

ss = 2
For mnt = 1 To ic

nv = ""
For ws = ss To ic
nv = nv + split(ws)
Next ws

ss = ss + 1
 Combo1.AddItem split(1) + ps
Combo2.AddItem nv

If ss > 0 And ss < ic Then ps = ps + split(ss - 1)


Next mnt
nv = ""
ss = 2
ps = ""
For mnt = 1 To ic
For ws = ss To ic
nv = nv + split(ws)
Next ws
ss = ss + 1
Combo2.AddItem split(1) + ps
Combo1.AddItem nv
If ss > 0 And ss < ic Then ps = ps + split(ss - 1)
nv = ""
Next mnt
List1.Clear
pis = 0
For pi = 1 To 15
If InStr(1, UCase(picg(pnum, pi)), "BACK", vbTextCompare) > 1 Then bs = 1
If InStr(1, UCase(picg(pnum, pi)), "FRONT", vbTextCompare) > 1 Then: fs = 1: pic = picg(pnum, pi): Image1.Picture = LoadPicture(pic): Label10.Visible = False
If e = 75 Then Label10.Visible = True: e = 0
If picg(pnum, pi) <> "" Then List1.AddItem picg(pnum, pi): nop = nop + 1
Next pi
GoTo skp
er:

If Err = 75 Then e = 75: r = MsgBox("There is a fault on file " + pic + " therfore I will not load the picture the program should carry on as normal.", vbOKOnly, "Error!"): Label10.Caption = "There is a fault on file " + pic + " therfore It will not load" Else Resume Next
Resume Next
skp:
End Sub

Private Sub Form_Terminate()
artg = Text4:  Text2 = Text4
titg = Text3:  Text1 = Text3
nocdg = Text5
comments(ind) = Text6
Text6 = ""

Form2.Hide
Form1.Enabled = True
Form1.Show
sf = 1
Form2.Enabled = False
End Sub

Private Sub Image1_Click()
prv = 1
Form3.Image1.Picture = Image1.Picture

Form2.Hide
 Form3.Show

End Sub

Private Sub List1_Click()
On Error GoTo er
Label10.Caption = "If there is no picture here then click on above list, if there is no pics in the list then ther are no covers in JPG, BMP or GIF Format in the current directory."

Label10.Visible = False
pic = picg(pnum, List1.ListIndex + 1)
Image1.Picture = LoadPicture(pic)
GoTo skp
er:

If Err = 75 Then e = 75: r = MsgBox("There is a fault on file " + pic + " therfore I will not load the picture the program should carry on as normal.", vbOKOnly, "Error!"): Label10.Caption = "There is a fault on file " + pic + " therfore It will not load" Else Resume Next
Label10.Visible = True
Resume Next
skp:
End Sub
Public Function SquishSpaces(ByVal strText As String) As String

    Const TWO_SPACES As String = "  "
   
    Dim intPos As Integer
    Dim strTemp As String
   
    intPos = InStr(1, strText, TWO_SPACES, vbBinaryCompare)
    Do While intPos > 0
        strTemp = LTrim$(Mid$(strText, intPos + 1))
        strText = Left$(strText, intPos) & strTemp
        intPos = InStr(1, strText, TWO_SPACES, vbBinaryCompare)
    Loop
   
   SquishSpaces = strText

End Function


Private Sub Option1_Click(Index As Integer)
fault(ind) = Index

End Sub
