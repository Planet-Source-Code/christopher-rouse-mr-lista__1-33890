VERSION 5.00
Begin VB.Form endSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   735
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "endSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   0
      Picture         =   "endSplash.frx":000C
      Top             =   0
      Width           =   6585
   End
End
Attribute VB_Name = "endSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End
End Sub

Private Sub Frame1_Click()
    Unload Me
End

End Sub

Private Sub lblCompanyProduct_Click()

End Sub

Private Sub Timer1_Timer()
End
End Sub
