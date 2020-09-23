VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2475
   LinkTopic       =   "Form2"
   ScaleHeight     =   885
   ScaleWidth      =   2475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Show AboutBox"
      Height          =   555
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2115
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  fAbout.Show 1
End Sub


Private Sub Form_Load()
  fAbout.Tag = "SPLASH"
  fAbout.Show
 ' fAbout.Refresh
 ' DoEvents
  d# = Timer
  Do While Timer - d < 3: DoEvents: Loop
  Unload fAbout
End Sub


