VERSION 5.00
Begin VB.Form fOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parameters"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   Icon            =   "fOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1800
      TabIndex        =   28
      Text            =   "Text6"
      Top             =   5460
      Width           =   5235
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      TabIndex        =   26
      Text            =   "Text5"
      Top             =   5160
      Width           =   5235
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   60
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   23
      Top             =   60
      Width           =   7095
      Begin VB.Image Image2 
         Height          =   735
         Index           =   2
         Left            =   6300
         Picture         =   "fOptions.frx":0982
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Open the VB project to translate and set up the options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   6075
      End
   End
   Begin VB.CommandButton Command5 
      Height          =   315
      Left            =   6660
      Picture         =   "fOptions.frx":0B6D
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Open VB Project"
      Top             =   1380
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1500
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   1380
      Width           =   5115
   End
   Begin VB.CommandButton Command4 
      Height          =   315
      Left            =   6660
      Picture         =   "fOptions.frx":0C6F
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Open VB Project"
      Top             =   1020
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1020
      Width           =   5115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6180
      TabIndex        =   16
      Top             =   5940
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   7035
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   "Single Page by item"
         Height          =   255
         Left            =   4320
         TabIndex        =   29
         Top             =   2340
         Width           =   2355
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   240
         Width           =   5715
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright Page"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Overview Page"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   1740
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Skip Private members"
         Height          =   195
         Left            =   4320
         TabIndex        =   9
         Top             =   1440
         Width           =   2355
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "What's New Page"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Generate HTML Help WorkShop Project and headers"
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   2700
         Value           =   1  'Checked
         Width           =   6015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   5295
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text4"
         Top             =   540
         Width           =   5295
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   6540
         Picture         =   "fOptions.frx":0D71
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   540
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "Skip Static controls"
         Height          =   195
         Left            =   4320
         TabIndex        =   3
         Top             =   1740
         Width           =   2355
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Registration Page"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   2
         Top             =   2340
         Width           =   2055
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "Include Empty items"
         Height          =   195
         Left            =   4320
         TabIndex        =   1
         Top             =   2040
         Width           =   2355
      End
      Begin VB.Label Label7 
         Caption         =   "Help Title"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Output Style"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label10 
         Caption         =   "Back. Image"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Width           =   1095
      End
   End
   Begin VB.Label Label4 
      Caption         =   "URL address"
      Height          =   255
      Left            =   180
      TabIndex        =   27
      Top             =   5460
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Help Author Name"
      Height          =   255
      Left            =   180
      TabIndex        =   25
      Top             =   5160
      Width           =   1515
   End
   Begin VB.Label Label6 
      Caption         =   "Destination Folder"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "VB Project"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1020
      Width           =   1095
   End
End
Attribute VB_Name = "fOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click(Index As Integer)
 Prj.AdFlag(Index) = Check1(Index) = 1
End Sub

Private Sub Check3_Click()
  Prj.GenHHP = Check3 = 1
End Sub

Private Sub Check4_Click()
     Prj.SkipStaticControls = Check4 = 1
End Sub

Private Sub Check6_Click()
SinglePage = Check6 = 1
End Sub

Private Sub Command1_Click()
 
     Prj.HTMLPath = Text2
     If Len(Prj.HTMLPath) = 0 Then
        MsgBox "Missing destination path", 16, "Type a valid name"
        Exit Sub
     End If
     
     On Error Resume Next
     MkDir Prj.HTMLPath
     If Err.Number <> 0 And Err.Number <> 75 Then
        MsgBox "Invalid destination path", 16, "wrong destination path name"
        Exit Sub
     End If
     On Error GoTo 0
 
     If Len(Text1) = 0 Then
        MsgBox "Missing VB Project", 16, "Type a valid name"
        Exit Sub
     End If
     Prj.ProjectName = Text3
     Prj.Author = Text5
     Prj.URL = Text6
 Unload Me
End Sub

Private Sub Check2_Click()
  Prj.PublicMBR = Check2 = 1
End Sub

Private Sub Check5_Click()
  Prj.IncludeEmptyItems = Check5 = 1
End Sub


Private Sub Combo1_Click()
  Prj.OutputMode = Combo1.ListIndex
  Check4.Enabled = Prj.OutputMode = 1
  Check2.Enabled = Prj.OutputMode = 0
    
End Sub



Private Sub Command2_Click()
  With FMain.CommonDialog1
    .Filter = "Image Files (*.gif,*.jpg)|*.gif;*.jpg"
    .FileName = ""
    .ShowOpen
    Text4 = .FileName
    
  End With
End Sub


Private Sub Command4_Click()
  With FMain.CommonDialog1
    .Filter = ".vbp Files|*.vbp"
    .FileName = ""
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    Text1 = .FileName
    Prj.ProjectFile = Text1
    
    OpenVBP Text1
    Text3 = Prj.ProjectName & " v" & Prj.ProjectVersion
  End With

End Sub

Sub OpenVBP(File As String)

 Dim St As String
 Dim C As cModule
 Dim cnt As Long
 Dim ClassName As String, FileName As String
 
 
 For i = 1 To Files.Count
    Files.Remove 1
 Next
 
 Prj.PathVBP = ExtractPathFromString(File)
 
 Open File For Input As #1
 
 Do Until EOF(1)
   cnt = cnt + 1
   Line Input #1, St
   Set C = New cModule
   If cnt = 1 And Mid(St, 1, 4) = "Type" Then
      Prj.ProjectType = Mid(St, 6, Len(St))
   End If
   If Mid(St, 1, 4) = "Name" Then
      Prj.ProjectName = Mid(St, 7, Len(St) - 7)
   End If
   
   If Mid(St, 1, 8) = "MajorVer" Then
     Prj.ProjectVersion = CStr(Val(Mid(St, 10, 10)))
   End If
   If Mid(St, 1, 8) = "MinorVer" Then Prj.ProjectVersion = Prj.ProjectVersion & "." & CStr(Val(Mid(St, 10, 10)))
   If Mid(St, 1, 11) = "RevisionVer" Then Prj.ProjectVersion = Prj.ProjectVersion & "." & CStr(Val(Mid(St, 13, 10)))
   

   If Mid(St, 1, 5) = "Class" Then
      C.ModuleType = "CLS"
      GetNames St, ClassName, FileName
      C.ClassName = ClassName
      C.FileName = FileName
   ElseIf Mid(St, 1, 6) = "Module" Then
      C.ModuleType = "BAS"
      GetNames St, ClassName, FileName
      C.ClassName = ClassName
      C.FileName = FileName
   ElseIf Mid(St, 1, 4) = "Form" Then
      C.ModuleType = "FRM"
      GetNames St, ClassName, FileName
      C.ClassName = ClassName
      C.FileName = FileName
   ElseIf Mid(St, 1, 11) = "UserControl" Then
      C.ModuleType = "CTL"
      GetNames St, ClassName, FileName
      C.ClassName = ClassName
      C.FileName = FileName
   End If
   
   If Len(C.ModuleType) > 0 Then
      C.Key = "R" & CStr(Files.Count + 1)
      C.Caption = C.ClassName
      Files.Add Item:=C, Key:=C.Key
   End If
   
'   If Files.Count > 5 And Demo Then
'      MsgBox "Unregistered version. Can't handle more than 5 files per project", 64, "Warning"
'      Close 1
'      Exit Sub
'   End If
   
 Loop
 
 Close 1
 
 
 For i = 1 To Files.Count
 
     Files(i).ResolveVariables
     Files(i).ResolveProcedures
     Files(i).ResolveObjects
    
 Next
 

End Sub




Private Sub Command5_Click()
    strOutputDir = BrowseForFolder(Me.hWnd, "Select a Folder to Write to", "c:\")
    Text2 = strOutputDir

End Sub

Private Sub Form_Load()
 Text1 = Prj.ProjectFile
 Text2 = Prj.HTMLPath
 Text3 = Prj.ProjectName
 Text4 = Prj.BackGroundImage
 Text5 = Prj.Author
 Text6 = Prj.URL
 
 If Len(Text5) = 0 Then Text5 = "VB Cad Geo Tools"
 If Len(Text6) = 0 Then Text6 = "http://www.sourcecode4free.com/cgt"
 
 Check5 = Abs(Prj.IncludeEmptyItems)
 Check2 = Abs(Prj.PublicMBR)
 Check4 = Abs(Prj.SkipStaticControls)
 Check3 = Abs(Prj.GenHHP)
 
 For i = 0 To 3
    Check1(i) = Abs(Prj.AdFlag(i))
 Next

 Combo1.AddItem "Programming Reference"
 Combo1.AddItem "Application Documentation"
 Combo1.ListIndex = 0

 If Len(Prj.ProjectFile) = 0 Then Command4_Click
End Sub



Private Sub Text1_Change()
 Prj.ProjectFile = Text1
End Sub



Private Sub Text4_Change()
 Prj.BackGroundImage = Text4
End Sub


