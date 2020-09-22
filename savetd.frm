VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTable1 
   Caption         =   "Table1"
   ClientHeight    =   9195
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   12660
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   12660
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3720
      TabIndex        =   31
      Top             =   7560
      Width           =   3975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Artist"
      Height          =   285
      Index           =   0
      Left            =   4800
      TabIndex        =   21
      Top             =   3960
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Back"
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   20
      Top             =   4275
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CD ID"
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   19
      Top             =   4605
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CDs"
      Height          =   285
      Index           =   3
      Left            =   4800
      TabIndex        =   18
      Top             =   4920
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Comments"
      Height          =   285
      Index           =   4
      Left            =   4800
      TabIndex        =   17
      Top             =   5235
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Fault"
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   16
      Top             =   5565
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Front"
      Height          =   285
      Index           =   6
      Left            =   4800
      TabIndex        =   15
      Top             =   5880
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "OK"
      Height          =   285
      Index           =   7
      Left            =   4800
      TabIndex        =   14
      Top             =   6195
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Title"
      Height          =   285
      Index           =   8
      Left            =   4800
      TabIndex        =   13
      Top             =   6525
      Width           =   3375
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   12660
      TabIndex        =   6
      Top             =   8595
      Width           =   12660
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1213
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   6240
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   3480
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   1213
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   7
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   12660
      TabIndex        =   0
      Top             =   8895
      Width           =   12660
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "savetd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "savetd.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "savetd.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "savetd.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label artist 
      Caption         =   "artist"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Back:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   4275
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CD ID:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   4605
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CDs:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Comments:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   26
      Top             =   5235
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fault:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   5565
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Front:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "OK:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   23
      Top             =   6195
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Title:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   6525
      Width           =   1815
   End
End
Attribute VB_Name = "frmTable1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Command1_Click()
Rem chkfields munst be 0 of 1
txtFields(0) = "modif"
chkFields(1) = "1"
txtFields(2) = "pm1444"
Call cmdNext_Click
End Sub

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=C:\My Documents\blankdatabses.mdb;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Artist,Back,[CD ID],CDs,Comments,Fault,Front,OK,Title from Table1", db, adOpenStatic, adLockOptimistic

  Set grdDataGrid.DataSource = adoPrimaryRS
 
Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Dim oCheck As CheckBox
  'Bind the check boxes to the data provider
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = adoPrimaryRS
  Next

  mbDataChanged = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  grdDataGrid.Height = Me.ScaleHeight - 30 - picButtons.Height - picStatBox.Height
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  adoPrimaryRS.MoveLast
  adoPrimaryRS.AddNew
  grdDataGrid.SetFocus

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS

  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub grdDataGrid_Click()

End Sub

