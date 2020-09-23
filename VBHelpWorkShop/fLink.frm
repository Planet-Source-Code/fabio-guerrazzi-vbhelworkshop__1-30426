VERSION 5.00
Begin VB.Form fLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Link"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   ".."
      Height          =   315
      Left            =   5280
      TabIndex        =   4
      Top             =   540
      Width           =   315
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   780
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   540
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   780
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Text"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   495
   End
End
Attribute VB_Name = "fLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Text1 = LinkedText
 Text2 = LinkedURL
End Sub

Private Sub Text1_Change()
  LinkedText = Text1
End Sub


Private Sub Text2_Change()
LinkedURL = Text2
End Sub


