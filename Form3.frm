VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Preview"
   ClientHeight    =   12060
   ClientLeft      =   990
   ClientTop       =   0
   ClientWidth     =   14415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12060
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   12135
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Image1_Click()
Form3.Hide
Form2.Show



sf = 1
End Sub
