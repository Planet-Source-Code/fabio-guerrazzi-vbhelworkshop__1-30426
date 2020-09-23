VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB Help WorkShop Wizard"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "FWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frames 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4875
      Index           =   3
      Left            =   60
      TabIndex        =   30
      Top             =   60
      Width           =   7275
      Begin VB.CommandButton Command6 
         Caption         =   ".."
         Height          =   315
         Left            =   6960
         TabIndex        =   62
         Top             =   2940
         Width           =   315
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   4320
         TabIndex        =   61
         Text            =   "Text8"
         Top             =   2940
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         Height          =   555
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Text            =   "FWizard.frx":0ABA
         Top             =   4320
         Width           =   5595
      End
      Begin VB.CommandButton Command5 
         Caption         =   ".."
         Height          =   315
         Left            =   2640
         TabIndex        =   56
         Top             =   2880
         Width           =   315
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   2880
         Width           =   1635
      End
      Begin VB.TextBox Text6 
         Height          =   615
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Text            =   "FWizard.frx":0AC0
         Top             =   3660
         Width           =   5595
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   51
         Text            =   "Text5"
         Top             =   3300
         Width           =   5595
      End
      Begin MSComctlLib.TreeView TV2 
         Height          =   1875
         Left            =   60
         TabIndex        =   49
         Top             =   960
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   3307
         _Version        =   393217
         Indentation     =   353
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   3
         Left            =   60
         ScaleHeight     =   825
         ScaleWidth      =   7185
         TabIndex        =   31
         Top             =   60
         Width           =   7215
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Document the items that will be shown on help pages"
            Height          =   375
            Index           =   2
            Left            =   420
            TabIndex        =   33
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Items Layout "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   60
            Width           =   5595
         End
         Begin VB.Image Image2 
            Height          =   735
            Index           =   2
            Left            =   6360
            Picture         =   "FWizard.frx":0AC6
            Top             =   60
            Width           =   735
         End
      End
      Begin VB.Label Label16 
         Caption         =   "ScreenShoot"
         Height          =   195
         Left            =   3120
         TabIndex        =   60
         Top             =   2940
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Remarks"
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Category"
         Height          =   255
         Left            =   60
         TabIndex        =   55
         Top             =   2940
         Width           =   795
      End
      Begin VB.Label Label13 
         Caption         =   "Description"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   3780
         Width           =   1155
      End
      Begin VB.Label Label12 
         Caption         =   "Caption/Item"
         Height          =   255
         Left            =   180
         TabIndex        =   50
         Top             =   3300
         Width           =   1275
      End
   End
   Begin VB.Frame Frames 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3315
      Index           =   4
      Left            =   5880
      TabIndex        =   45
      Top             =   1440
      Width           =   6615
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   4
         Left            =   60
         ScaleHeight     =   825
         ScaleWidth      =   7185
         TabIndex        =   46
         Top             =   60
         Width           =   7215
         Begin VB.Image Image2 
            Height          =   735
            Index           =   3
            Left            =   6360
            Picture         =   "FWizard.frx":0CB1
            Top             =   60
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "VB Help WorkShop has been completed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   48
            Top             =   60
            Width           =   5595
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Start Microsoft HTML Help Workshop and compile the hhp file to obtain a .chm file"
            Height          =   375
            Index           =   3
            Left            =   420
            TabIndex        =   47
            Top             =   360
            Width           =   5055
         End
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   315
      Left            =   60
      Picture         =   "FWizard.frx":0E9C
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5160
      Width           =   375
   End
   Begin VB.Frame Frames 
      BorderStyle     =   0  'None
      Height          =   4695
      Index           =   2
      Left            =   -660
      TabIndex        =   16
      Top             =   -180
      Width           =   7215
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   60
         TabIndex        =   23
         Top             =   1620
         Width           =   7035
         Begin VB.CheckBox Check5 
            Alignment       =   1  'Right Justify
            Caption         =   "Include Empty items"
            Height          =   195
            Left            =   4320
            TabIndex        =   59
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Registration Page"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   44
            Top             =   2340
            Width           =   2055
         End
         Begin VB.CheckBox Check4 
            Alignment       =   1  'Right Justify
            Caption         =   "Skip Static controls"
            Height          =   195
            Left            =   4320
            TabIndex        =   43
            Top             =   1740
            Width           =   2055
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   6360
            Picture         =   "FWizard.frx":0FE6
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   540
            Width           =   375
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   780
            TabIndex        =   38
            Text            =   "Text4"
            Top             =   540
            Width           =   5535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   960
            Width           =   4875
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Generate HTML Help WorkShop Project and headers"
            Height          =   315
            Left            =   240
            TabIndex        =   34
            Top             =   2640
            Value           =   1  'Checked
            Width           =   6015
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "What's New Page"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   29
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CheckBox Check2 
            Alignment       =   1  'Right Justify
            Caption         =   "Skip Private members"
            Height          =   195
            Left            =   4320
            TabIndex        =   28
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Overview Page"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   1740
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "Copyright Page"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   780
            TabIndex        =   25
            Text            =   "Text3"
            Top             =   240
            Width           =   6135
         End
         Begin VB.Label Label10 
            Caption         =   "Back. Image"
            Height          =   555
            Left            =   120
            TabIndex        =   39
            Top             =   540
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Output Style"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Title"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1860
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   1020
         Width           =   5175
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   1500
         Picture         =   "FWizard.frx":10E8
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1020
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   2
         Left            =   60
         ScaleHeight     =   825
         ScaleWidth      =   7185
         TabIndex        =   17
         Top             =   60
         Width           =   7215
         Begin VB.Image Image2 
            Height          =   735
            Index           =   1
            Left            =   6360
            Picture         =   "FWizard.frx":11EA
            Top             =   60
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Output Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   60
            Width           =   2535
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Choose your favourite way to build the pages"
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   18
            Top             =   420
            Width           =   4755
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Destination Folder"
         Height          =   195
         Left            =   60
         TabIndex        =   22
         Top             =   1080
         Width           =   1395
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3180
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   60
      TabIndex        =   11
      Top             =   4920
      Width           =   7275
   End
   Begin VB.Frame Frames 
      BorderStyle     =   0  'None
      Caption         =   "1"
      Height          =   4875
      Index           =   1
      Left            =   3420
      TabIndex        =   8
      Top             =   120
      Width           =   7395
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   240
         Picture         =   "FWizard.frx":13D5
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1140
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   660
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1140
         Width           =   6435
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   1
         Left            =   60
         ScaleHeight     =   825
         ScaleWidth      =   7185
         TabIndex        =   9
         Top             =   60
         Width           =   7215
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Select the VB project to translate as html pages"
            Height          =   315
            Index           =   0
            Left            =   420
            TabIndex        =   12
            Top             =   360
            Width           =   4755
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "VB Project"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   60
            Width           =   2535
         End
         Begin VB.Image Image2 
            Height          =   735
            Index           =   0
            Left            =   6360
            Picture         =   "FWizard.frx":14D7
            Top             =   60
            Width           =   735
         End
      End
      Begin MSComctlLib.TreeView TV 
         Height          =   3615
         Left            =   300
         TabIndex        =   15
         Top             =   1680
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   6376
         _Version        =   393217
         Indentation     =   529
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   6120
      TabIndex        =   6
      Top             =   5100
      Width           =   1125
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   345
      Left            =   4860
      TabIndex        =   5
      Top             =   5100
      Width           =   1125
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3720
      TabIndex        =   4
      Top             =   5100
      Width           =   1125
   End
   Begin VB.Frame Frames 
      BorderStyle     =   0  'None
      Height          =   4875
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7395
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4695
         Index           =   0
         Left            =   2520
         ScaleHeight     =   4665
         ScaleWidth      =   4725
         TabIndex        =   1
         Top             =   60
         Width           =   4755
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "click on Cancel to exit VB Help WorkShop"
            ForeColor       =   &H80000010&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   4380
            Width           =   4455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"FWizard.frx":16C2
            Height          =   1275
            Left            =   120
            TabIndex        =   3
            Top             =   960
            Width           =   4335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Welcome to VB Help WorkShop wizard"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   4275
         End
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   180
         TabIndex        =   42
         Top             =   4500
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   4710
         Left            =   60
         Picture         =   "FWizard.frx":1751
         Top             =   60
         Width           =   2460
      End
   End
   Begin VB.Label Label9 
      Caption         =   "VBcgt ©2001"
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   600
      TabIndex        =   37
      Top             =   5220
      Width           =   2835
   End
End
Attribute VB_Name = "FWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurFrame As Long
Const NF = 4

 
Sub BuildAdditionalPages()

Dim File As String


' ********** crea le pagine aggiuntive

 If Check1(0) = 1 Then
    File = HTMLPath & "\CopyRight.htm"
    AddFiles(1) = File
    TPL$ = App.Path & "\templates\legal.txt"
    Open File For Output As 1
    Open TPL For Input As 2
      WriteHeader 1
      Do Until EOF(2)
         Line Input #2, St
         Print #1, St
      Loop
     PrintFooter 1
    Close 1
    Close 2
 End If

 If Check1(2) = 1 Then
    File = HTMLPath & "\WhatsNew.htm"
    AddFiles(2) = File
    TPL$ = App.Path & "\templates\WhatsNew.txt"
    Open File For Output As 1
    Open TPL For Input As 2
      WriteHeader 1
      Do Until EOF(2)
         Line Input #2, St
         Print #1, St
      Loop
     PrintFooter 1
    Close 1
    Close 2
 End If

 
 

' Crea i files di progetto per Microsoft HTML WorkShop


 If Check3 = 1 Then ' Generate HTML WorkShop files
   File = HTMLPath & "\" & ProjectName & ".hhp"
   Open File For Output As 1
    Print #1, "[Options]"
    Print #1, "Compatibility = 1.1"
    Print #1, "Compiled File = " & ProjectName & ".chm"
    Print #1, "Contents file=" & HTMLPath & "\Table_of_Contents.hhc"
    Print #1, "Default topic=" & HTMLPath & "\Index.htm"
    Print #1, "Display compile progress=No"
    'Print #1,"Language=0x410 Italiano (Italia)"

    Print #1, "[Files]"
    For i = 1 To Files.Count
            Print #1, HTMLPath & "\" & TrimExt(Files(i).ClassName) & ".htm"
    Next

    Close 1
    
' Table of Contents
   File = HTMLPath & "\Table_of_Contents.hhc"
   Open File For Output As 1
     Print #1, "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">"
     Print #1, "<HTML>"
     Print #1, "<HEAD>"
     Print #1, "<meta name=""GENERATOR"" content=""Mandix&reg; VB Help Workshop 1.0"">"
     Print #1, "<!-- Sitemap 1.0 -->"
     Print #1, "</HEAD><BODY>"
     Print #1, "<OBJECT type=""text/site properties"">"
     Print #1, "    <param name=""ImageType"" value=""Folder"">"
     Print #1, "</OBJECT>"
     Print #1, "<UL>"
     Print #1, "    <LI> <OBJECT type=""text/sitemap"">"
     Print #1, "        <param name=""Name"" value=""" & ProjectName & " Guide"">"
     Print #1, "        <param name=""ImageNumber"" value=""1"">"
     Print #1, "        </OBJECT>"
     Print #1, "    <UL>"
     Print #1, "        <LI> <OBJECT type=""text/sitemap"">"
     Print #1, "            <param name=""Name"" value=""Index"">"
     Print #1, "            <param name=""Local"" value=""Index.htm"">"
     Print #1, "            </OBJECT>"
     Print #1, "    </UL>"
     
     For i = 1 To 5
       If Len(AddFiles(i)) > 0 Then
             FL$ = ExtractFileFromString(AddFiles(i))
             TT$ = EstraiNomeFile(FL)
             Print #1, "    <UL>"
             Print #1, "        <LI> <OBJECT type=""text/sitemap"">"
             Print #1, "            <param name=""Name"" value=""" & TT & """>"
             Print #1, "            <param name=""Local"" value=""" & FL & """>"
             Print #1, "            </OBJECT>"
             Print #1, "    </UL>"
       End If
     Next
     
     ok = False
     For i = 1 To Files.Count
         If Files(i).ModuleType = "CTL" Then
            If Not ok Then
                Print #1, "    <LI> <OBJECT type=""text/sitemap"">"
                Print #1, "        <param name=""Name"" value=""Controls"">"
                Print #1, "        <param name=""ImageNumber"" value=""1"">"
                Print #1, "        </OBJECT>"
                ok = True
            End If
            Print #1, "    <UL>"
            Print #1, "        <LI> <OBJECT type=""text/sitemap"">"
            Print #1, "            <param name=""Name"" value=""" & Files(i).ClassName & """>"
            Print #1, "            <param name=""Local"" value=""" & TrimExt(Files(i).ClassName) & ".htm"">"
            Print #1, "            </OBJECT>"
            Print #1, "    </UL>"
         End If
     Next
     
     
     ok = False
     For i = 1 To Files.Count
         If Files(i).ModuleType = "FRM" Then
            If Not ok Then
                Print #1, "    <LI> <OBJECT type=""text/sitemap"">"
                Print #1, "        <param name=""Name"" value=""Forms"">"
                Print #1, "        <param name=""ImageNumber"" value=""1"">"
                Print #1, "        </OBJECT>"
                ok = True
            End If
            Print #1, "    <UL>"
            Print #1, "        <LI> <OBJECT type=""text/sitemap"">"
            Print #1, "            <param name=""Name"" value=""" & Files(i).ClassName & """>"
            Print #1, "            <param name=""Local"" value=""" & TrimExt(Files(i).ClassName) & ".htm"">"
            Print #1, "            </OBJECT>"
            Print #1, "    </UL>"
            'Print #1, HTMLPath & "\" & TrimExt(Files(i).ClassName) & ".htm"
         End If
     Next
   
   If OutputMode = 1 Then GoTo naa
     
     ok = False
     For i = 1 To Files.Count
         If Files(i).ModuleType = "CLS" Then
            If Not ok Then
                Print #1, "    <LI> <OBJECT type=""text/sitemap"">"
                Print #1, "        <param name=""Name"" value=""Objects"">"
                Print #1, "        <param name=""ImageNumber"" value=""1"">"
                Print #1, "        </OBJECT>"
                ok = True
            End If
            Print #1, "    <UL>"
            Print #1, "        <LI> <OBJECT type=""text/sitemap"">"
            Print #1, "            <param name=""Name"" value=""" & Files(i).ClassName & """>"
            Print #1, "            <param name=""Local"" value=""" & TrimExt(Files(i).ClassName) & ".htm"">"
            Print #1, "            </OBJECT>"
            Print #1, "    </UL>"
         End If
     Next
     
     ok = False
     For i = 1 To Files.Count
         If Files(i).ModuleType = "BAS" Then
            If Not ok Then
                Print #1, "    <LI> <OBJECT type=""text/sitemap"">"
                Print #1, "        <param name=""Name"" value=""Modules"">"
                Print #1, "        <param name=""ImageNumber"" value=""1"">"
                Print #1, "        </OBJECT>"
                ok = True
            End If
            Print #1, "    <UL>"
            Print #1, "        <LI> <OBJECT type=""text/sitemap"">"
            Print #1, "            <param name=""Name"" value=""" & Files(i).ClassName & """>"
            Print #1, "            <param name=""Local"" value=""" & TrimExt(Files(i).ClassName) & ".htm"">"
            Print #1, "            </OBJECT>"
            Print #1, "    </UL>"
         End If
     Next
     
naa:

     Print #1, "</UL>"
     Print #1, "</BODY></HTML>"
    
   Close 1
    
    
 End If
 
 
End Sub

Sub Convert()
  Dim Indexfile As String
  Dim ok As Boolean
  Indexfile = HTMLPath & "\Index.htm"
  Open Indexfile For Output As 1
  
  
  If Len(BackGroundImage) > 0 Then
     fimg$ = HTMLPath & "\" & ExtractFileFromString(BackGroundImage)
     If Not FileExists(fimg) Then
        FileCopy BackGroundImage, fimg
     End If
  End If
  
  WriteHeader 1
  
  ok = False
  For i = 1 To Files.Count
      If Files(i).ModuleType = "CTL" And OutputMode = 0 Then
         If Not ok Then
             Print #1, "<BR><strong><FONT SIZE=""3"" COLOR=""#000040"">Controls</strong></FONT>"
             Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
             ok = True
         End If
         Print #1, "<A HREF=""" & TrimExt(Files(i).ClassName) & ".htm"">" & TrimExt(Files(i).ClassName) & "</A>"
         Print #1, "<BR>"
      End If
  Next
  
  
  ok = False
  For i = 1 To Files.Count
      If Files(i).ModuleType = "FRM" Then
         If Not ok Then
           Print #1, "<BR><strong><FONT SIZE=""3"" COLOR=""#000040"">Forms</strong></FONT>"
           Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
           ok = True
         End If
         Print #1, "<A HREF=""" & TrimExt(Files(i).ClassName) & ".htm"">" & TrimExt(Files(i).ClassName) & "</A>"
         Print #1, "<BR>"
      End If
  Next
  
  If OutputMode > 0 Then GoTo NoProgr ' if App Doc only skips the next statements
  
  ok = False
  For i = 1 To Files.Count
      If Files(i).ModuleType = "CLS" Then
         If Not ok Then
            Print #1, "<BR><strong><FONT SIZE=""3"" COLOR=""#000040"">Object Classes</strong></FONT>"
            Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
            ok = True
         End If
         Print #1, "<A HREF=""" & TrimExt(Files(i).ClassName) & ".htm"">" & TrimExt(Files(i).ClassName) & "</A>"
         Print #1, "<BR>"
      End If
  Next
  
  
  ok = False
  For i = 1 To Files.Count
      If Files(i).ModuleType = "BAS" Then
         If Not ok Then
             Print #1, "<BR><strong><FONT SIZE=""3"" COLOR=""#000040"">Modules</strong></FONT>"
             Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
             ok = True
         End If
         Print #1, "<A HREF=""" & TrimExt(Files(i).ClassName) & ".htm"">" & TrimExt(Files(i).ClassName) & "</A>"
         Print #1, "<BR>"
      End If
  Next
  
NoProgr:
  
  PrintFooter 1
  
  Close 1
  
  
' Save modules to html pages
  For i = 1 To Files.Count
      If OutputMode = 0 Then
         Files(i).ExportToHtml
      ElseIf Files(i).ModuleType = "FRM" Then
         Files(i).WriteDoc
      End If
  Next

  BuildAdditionalPages

End Sub

Sub FillCats()
   On Error Resume Next
   Combo2.Clear
   Combo2.AddItem "<Global Topic>"
   For i = 1 To nCat
      Combo2.AddItem Cats(i)
   Next
   Combo2.ListIndex = 0
End Sub

Sub FillTV2()
 Dim itmx As Node, i As Long, KD As String, Dx As String
 Dim ok As Boolean
 
 TV2.Nodes.Clear
 
Set itmx = TV2.Nodes.Add(, , "PRJ0", "Project: " & ProjectName & "[" & ProjectType & "]")
If OutputMode = 1 Then
 ok = False
 For i = 1 To Files.Count
     If Files(i).ModuleType = "FRM" Then
        If Not ok Then Set itmx = TV2.Nodes.Add("PRJ0", tvwChild, "FRM", "Forms")
        Set itmx = TV2.Nodes.Add("FRM", tvwChild, Files(i).Key, Files(i).ClassName)
        For j = 1 To Files(i).Objects.Count
           Files(i).Objects(j).AddToTree TV2, Files(i).Key
        Next
        ok = True
     End If
 Next
Else
 On Error Resume Next
 For i = 1 To Files.Count
    GoSub GetType
    Set itmx = TV2.Nodes.Add("PRJ0", tvwChild, Files(i).ModuleType, Dx)
 Next
 
 For i = 1 To Files.Count
        GoSub GetType
        Set itmx = TV2.Nodes.Add(KD, tvwChild, Files(i).Key, Files(i).ClassName)
        For j = 1 To Files(i).Procs.Count
            Set itmx = TV2.Nodes.Add(Files(i).Key, tvwChild, Files(i).Procs(j).Key, Files(i).Procs(j).FullName)
        Next
 Next
End If

Exit Sub
 
 
GetType:
    KD$ = Files(i).ModuleType
    Select Case KD
     Case "FRM": Dx$ = "Forms"
     Case "BAS": Dx$ = "Modules"
     Case "CLS": Dx$ = "Object Classes"
     Case "CTL": Dx$ = "Object Controls"
    End Select

Return
 
 
 
 '======================
 Set itmx = TV2.Nodes.Add("PRJ0", tvwChild, "FRM", "Forms")
 Set itmx = TV2.Nodes.Add("PRJ0", tvwChild, "BAS", "Modules")
 Set itmx = TV2.Nodes.Add("PRJ0", tvwChild, "CLS", "Classes")
 Set itmx = TV2.Nodes.Add("PRJ0", tvwChild, "CTL", "User Controls")
 
 For i = 1 To Files.Count
 
   Select Case Files(i).ModuleType
     Case "BAS": Root$ = "BAS"
     Case "CLS": Root$ = "CLS"
     Case "FRM": Root$ = "FRM"
     Case "CTL": Root$ = "CTL"
   End Select
       
     Set itmx = TV2.Nodes.Add(Root, tvwChild, Files(i).Key, Files(i).ClassName)
     
     Set itmx = TV2.Nodes.Add(Files(i).Key, tvwChild, "V" & Files(i).Key, "Declarations")
     Set itmx = TV2.Nodes.Add("V" & Files(i).Key, tvwChild, "U1" & Files(i).Key, "User Types")
     Set itmx = TV2.Nodes.Add("V" & Files(i).Key, tvwChild, "U2" & Files(i).Key, "Public")
     Set itmx = TV2.Nodes.Add("V" & Files(i).Key, tvwChild, "U3" & Files(i).Key, "Private")
     Set itmx = TV2.Nodes.Add(Files(i).Key, tvwChild, "P" & Files(i).Key, "Public Procedures")
     Set itmx = TV2.Nodes.Add(Files(i).Key, tvwChild, "R" & Files(i).Key, "Private Procedures")
     
     For j = 1 To Files(i).UTypes.Count
        With Files(i).UTypes(j)
            Root$ = "U1" & Files(i).Key
            Key$ = Root$ & Format(j, "00")
            Set itmx = TV2.Nodes.Add(Root$, tvwChild, Key, .Name)
            Root$ = Key
            For b = 1 To .Vars.Count
                Key$ = Root & Format(b, "00")
                Var$ = .Vars(b).Name & " As " & .Vars(b).varType
               Set itmx = TV2.Nodes.Add(Root, tvwChild, Key, Var)
           Next
       End With
     Next
     
     
     For j = 1 To Files(i).Vars.Count
        With Files(i).Vars(j)
            If .Mode = "Public" Then
                Root$ = "U2" & Files(i).Key
            Else
                Root$ = "U3" & Files(i).Key
            End If
            Key$ = Root$ & Format(j, "00")
            Var$ = .Name & " As " & .varType
            Set itmx = TV2.Nodes.Add(Root$, tvwChild, Key, Var)
        End With
     Next
     
     
     
     For j = 1 To Files(i).Procs.Count
        With Files(i).Procs(j)
          If .IsPublic Then
              Root$ = "P" & Files(i).Key
          Else
              Root$ = "R" & Files(i).Key
          End If
            Key$ = Root$ & Format(j, "00")
            Set itmx = TV2.Nodes.Add(Root$, tvwChild, Key, .FullName)
        End With
     Next
    
 Next

End Sub

Sub ShowFrame()
 Dim i&
 For i = 0 To NF
   Frames(i).Visible = False
 Next
  Frames(CurFrame).Visible = True
End Sub


'<%*************************************************************
'                                                              *
'<Company>: VB CAD/Geo Tools
'<Author>:  Fabio Guerrazzi
'
'<Version>: 1.0.2
'<Date>:    06/08/01 10.14.10
'
'<Method>:  Check2_Click
'
'<Description>:
'//
'hhhhhhhhh
'//
'<Parameters>:
'//
'kkkk
'
'//
'                                                              *
'*************************************************************%>

Private Sub Check2_Click()
  PublicMBR = Check2 = 1
End Sub

Private Sub Check5_Click()
  IncludeEmptyItems = Check5 = 1
End Sub

Private Sub cmdBack_Click()
  CurFrame = CurFrame - 1
  If CurFrame = 0 Then cmdBack.Enabled = False
  ShowFrame
End Sub

Private Sub cmdCancel_Click()
  End
End Sub


Private Sub cmdNext_Click()
  
  If CurFrame = 1 And Len(ProjectName) = 0 Then
      MsgBox "Project required", 16
      Exit Sub
  End If
  
  If CurFrame = NF - 1 Then
     If Len(Text2) = 0 Then
        MsgBox "Destination Path required", 16
        Exit Sub
     End If
     If Len(Text3) = 0 Then
        MsgBox "Project Name required", 16
        Exit Sub
     End If
  End If
  
  If CurFrame = 2 Then '
     FillCats
     BackGroundImage = Text4
     SkipStaticControls = Check4 = 1
     HTMLPath = Text2
     On Error Resume Next
     MkDir HTMLPath
     If Err.Number <> 0 And Err.Number <> 75 Then
        MsgBox "Invalid path", 16, "wrong destination path name"
        Exit Sub
     End If
     On Error GoTo 0
     ProjectName = Text3
     FillTV2
  End If
  
  CurFrame = CurFrame + 1
  cmdBack.Enabled = True
  
  If CurFrame = NF Then
     Convert
     cmdNext.Caption = "&Finish"
  End If
  
  If CurFrame = NF + 1 Then
     End
  End If
  
  ShowFrame
End Sub


Private Sub Combo1_Click()
  OutputMode = Combo1.ListIndex
  Check4.Enabled = OutputMode = 1
  Check2.Enabled = OutputMode = 0
    
End Sub


Private Sub Combo2_Click()
  CurObj.HelpTopic = Combo2.List(Combo2.ListIndex)
End Sub


Private Sub Command1_Click()
  With CommonDialog1
    .Filter = ".vbp Files|*.vbp"
    .FileName = ""
    .ShowOpen
    Text1 = .FileName
    If Len(Text1) = 0 Then Exit Sub
    OpenVBP Text1
    Text3 = ProjectName & " v" & ProjectVersion
  End With
End Sub

Sub OpenVBP(File As String)

 Dim St As String
 Dim C As cModule
 Dim cnt As Long
 Dim ClassName As String, FileName As String
 
 PathVBP = ExtractPathFromString(File)
 
 Open File For Input As #1
 
 Do Until EOF(1)
   cnt = cnt + 1
   Line Input #1, St
   Set C = New cModule
   If cnt = 1 And Mid(St, 1, 4) = "Type" Then
      ProjectType = Mid(St, 6, Len(St))
   End If
   If Mid(St, 1, 4) = "Name" Then
      ProjectName = Mid(St, 7, Len(St) - 7)
   End If
   
   If Mid(St, 1, 8) = "MajorVer" Then
     ProjectVersion = CStr(Val(Mid(St, 10, 10)))
   End If
   If Mid(St, 1, 8) = "MinorVer" Then ProjectVersion = ProjectVersion & "." & CStr(Val(Mid(St, 10, 10)))
   If Mid(St, 1, 11) = "RevisionVer" Then ProjectVersion = ProjectVersion & "." & CStr(Val(Mid(St, 13, 10)))
   
      

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
   
   If Files.Count > 5 And Demo Then
      MsgBox "Unregistered version. Can't handle more than 5 files per project", 64, "Warning"
      Close 1
      Exit Sub
   End If
   
 Loop
 
 Close 1
 
 Dim itmx As Node
 TV.Nodes.Clear
 
     Set itmx = TV.Nodes.Add(, , "PRJ0", "Project: " & ProjectName & "[" & ProjectType & "]")
     Set itmx = TV.Nodes.Add("PRJ0", tvwChild, "FRM", "Forms")
     Set itmx = TV.Nodes.Add("PRJ0", tvwChild, "BAS", "Modules")
     Set itmx = TV.Nodes.Add("PRJ0", tvwChild, "CLS", "Classes")
     Set itmx = TV.Nodes.Add("PRJ0", tvwChild, "CTL", "User Controls")
 
 For i = 1 To Files.Count
 
   Select Case Files(i).ModuleType
    Case "BAS": Root$ = "BAS"
    Case "CLS": Root$ = "CLS"
    Case "FRM": Root$ = "FRM"
    Case "CTL": Root$ = "CTL"
   End Select
   
     
     Set itmx = TV.Nodes.Add(Root, tvwChild, Files(i).Key, Files(i).ClassName)
     
    ' Set itmx = TV.Nodes.Add(Files(i).Key, tvwChild, "V" & Files(i).Key, "Declarations")
    ' Set itmx = TV.Nodes.Add("V" & Files(i).Key, tvwChild, "U1" & Files(i).Key, "User Types")
    ' Set itmx = TV.Nodes.Add("V" & Files(i).Key, tvwChild, "U2" & Files(i).Key, "Public")
    ' Set itmx = TV.Nodes.Add("V" & Files(i).Key, tvwChild, "U3" & Files(i).Key, "Private")
   '  Set itmx = TV.Nodes.Add(Files(i).Key, tvwChild, "P" & Files(i).Key, "Public Procedures")
   '  Set itmx = TV.Nodes.Add(Files(i).Key, tvwChild, "R" & Files(i).Key, "Private Procedures")
     
     Files(i).ResolveVariables
     
    ' For j = 1 To Files(i).UTypes.Count
    '    With Files(i).UTypes(j)
    '        Root$ = "U1" & Files(i).Key
    '        Key$ = Root$ & Format(j, "00")
    '        Set itmx = TV.Nodes.Add(Root$, tvwChild, Key, .Name)
    '        Root$ = Key
    '        For b = 1 To .Vars.Count
    '            Key$ = Root & Format(b, "00")
    '            Var$ = .Vars(b).Name & " As " & .Vars(b).varType
    '           Set itmx = TV.Nodes.Add(Root, tvwChild, Key, Var)
    '       Next
    '   End With
    ' Next
     
     
   '  For j = 1 To Files(i).Vars.Count
   '     With Files(i).Vars(j)
   '         If .Mode = "Public" Then
   '             Root$ = "U2" & Files(i).Key
   '         Else
   '             Root$ = "U3" & Files(i).Key
   ''         End If
   '         Key$ = Root$ & Format(j, "00")
   '         Var$ = .Name & " As " & .varType
   '         Set itmx = TV.Nodes.Add(Root$, tvwChild, Key, Var)
   '     End With
   '  Next
     
     
     
     Files(i).ResolveProcedures
   '  For j = 1 To Files(i).Procs.Count
   '     With Files(i).Procs(j)
   '       If .IsPublic Then
   '           Root$ = "P" & Files(i).Key
   '       Else
   '           Root$ = "R" & Files(i).Key
   '       End If
   '         Key$ = Root$ & Format(j, "00")
   '         Set itmx = TV.Nodes.Add(Root$, tvwChild, Key, .FullName)
   '     End With
   '  Next
    Files(i).ResolveObjects
    
 Next

End Sub


Private Sub Command2_Click()
  With CommonDialog1
    .Filter = "Image Files (*.gif,*.jpg)|*.gif;*.jpg"
    .FileName = ""
    .ShowOpen
    Text4 = .FileName
  End With
End Sub

Private Sub Command3_Click()
    strOutputDir = BrowseForFolder(Me.hWnd, "Select a Folder to Write to", "c:\")
    Text2 = strOutputDir
End Sub

Private Sub Command4_Click()
  fAbout.Show 1
End Sub

Private Sub Command5_Click()
  fCat.Show 1
  FillCats
End Sub


Private Sub Command6_Click()
  With CommonDialog1
    .Filter = "Image Files (*.gif,*.jpg)|*.gif;*.jpg"
    .FileName = ""
    .ShowOpen
    Text8 = .FileName
  End With

End Sub

Private Sub Form_Load()
 Set CurObj = New cObj
 Text1 = ""
 Text2 = App.Path
 Text4 = ""
 Text5 = ""
 Text6 = ""
 Text7 = ""
 Text8 = ""
 
 License = GetSetting("VBHW", "Settings", "License")
 Demo = License <> "GBD-A65-7YU-340-54DB"
 If Demo Then Label11 = "Unregistered" Else Label11 = ""
 
 CurFrame = 0
 
 Combo1.AddItem "Programming Reference"
 Combo1.AddItem "Application Documentation"
 Combo1.ListIndex = 0
 
 For i = 1 To NF
   Frames(i).Caption = ""
   Frames(i).Move Frames(0).Left, Frames(0).Top, Frames(0).Width, Frames(0).Height
 Next
 ShowFrame
 
End Sub


Private Sub Text5_Change()
  CurObj.Caption = Text5
End Sub

Private Sub Text6_Change()
 CurObj.HelpDex = Text6
End Sub

Private Sub Text7_Change()
  CurObj.HelpRemarks = Text7
End Sub


Private Sub Text8_Change()
  CurObj.HelpScreenShoot = Text8
End Sub

Private Sub TV2_NodeClick(ByVal Node As MSComctlLib.Node)
  Dim Obj As Object
  Dim i As Long, j As Long, Found As Boolean
  
  For i = 1 To Files.Count
      If Files(i).Key = Node.Key Then
         Set Obj = Files(i)
         Found = True
         Exit For
      Else
        For j = 1 To Files(i).Objects.Count
            Set Obj = Files(i).Objects(j).GetObj(Node.Key)
            If Not Obj Is Nothing Then
                Found = True
                GoTo Via
            End If
        Next
        For j = 1 To Files(i).Procs.Count
            If Files(i).Procs(j).Key = Node.Key Then
              Set Obj = Files(i).Procs(j)
              Found = True
              GoTo Via
            End If
        Next
     End If
  Next
  
Via:

If Found Then
   Set CurObj = Obj
   Text5 = CurObj.Caption
   Text6 = CurObj.HelpDex
   Text7 = CurObj.HelpRemarks
   Text8 = CurObj.HelpScreenShoot
End If

End Sub


