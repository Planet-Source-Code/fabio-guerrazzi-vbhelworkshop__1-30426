VERSION 5.00
Begin VB.Form fAbout 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3165
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5625
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   180
   End
   Begin VB.CommandButton Command2 
      Caption         =   "More"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2640
      Width           =   915
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1140
      TabIndex        =   8
      Top             =   1500
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   5460
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   4740
      Picture         =   "About.frx":014A
      Top             =   1140
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Producer, dealer and supports:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   1860
      Width           =   5235
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   300
      Picture         =   "About.frx":0C04
      Top             =   1020
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "email fabiog2@libero.it"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   4
      Top             =   2700
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VB Help WorkShop v1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VB CAD/Geo Tools © 2001"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   780
      Width           =   3675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.sourcecode4free.com/cgt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   2520
      Width           =   3015
   End
End
Attribute VB_Name = "fAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Command2_Click()
 Call Shell("Notepad readme.txt", 1)
End Sub

Private Sub Form_Load()
 If Demo Then
   Label3 = "Unregistered"
 Else
   Label3 = "License Key=" & GetSetting("VBHW", "Settings", "License")
 End If
 Label5 = "v" & App.Major & "." & App.Revision & " " & App.Comments
End Sub


Private Sub Timer1_Timer()
  If Tag = "SPLASH" Then
     Command2.Visible = False
     Command1.Visible = False
  End If
  Timer1.Enabled = False
End Sub


