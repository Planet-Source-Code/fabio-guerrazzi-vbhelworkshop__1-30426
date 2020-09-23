VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   2235
      Left            =   5460
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "Form1.frx":0000
      Top             =   6360
      Width           =   5175
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   7635
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   13467
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   ".."
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   600
      Width           =   315
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   660
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Convert"
      Height          =   435
      Left            =   5520
      TabIndex        =   2
      Top             =   180
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   660
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   180
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   ".."
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   180
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "Dest Folder"
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Project"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   180
      Width           =   555
   End
   Begin VB.Menu file_m 
      Caption         =   "File"
      Begin VB.Menu New_i 
         Caption         =   "New"
      End
      Begin VB.Menu Open_i 
         Caption         =   "Open"
      End
      Begin VB.Menu Save_i 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu Tools_m 
      Caption         =   "Tools"
      Begin VB.Menu Indexpage_i 
         Caption         =   "generate Index Page"
      End
      Begin VB.Menu ShowCode_i 
         Caption         =   "Show Code"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub GetProcedure(idx As Long)
  Dim St As String
  Dim Ar() As String, Count As Long
  Dim C As GroupVars
  
  Open Files(idx).File For Input As #1
  Do Until EOF(1)
     Line Input #1, St
     Ar = Split(St, " ")
     On Error Resume Next
     Count = 0
     Count = UBound(Ar)
     If Count > 0 Then
        
        If Ar(0) = "Public" Or _
               Ar(0) = "Global" Then
           GoSub GetVar
           C.Mode = "PUB"
        ElseIf Ar(0) = "Private" Or _
               Ar(0) = "Dim" Then
           GoSub GetVar
           C.Mode = "LOC"
        End If
        
     End If
  Loop
  
  Close 1

Exit Sub

GetVar:
           If Ar(1) = "Enum" Then
           Else
              C.vName = Ar(1)
              If Count < 2 Then
                 C.vType = GetFixedType(C.vName)
                 If C.vType = 0 Then C.vType = "Variant"
              Else
                 C.vType = Ar(3)
              End If
           End If
Return
End Sub

Sub OpenVBP(File As String)

 Dim St As String
 Dim C As cModule
 Dim Cnt As Long
 
 Open File For Input As #1
 
 Do Until EOF(1)
   Cnt = Cnt + 1
   Line Input #1, St
   Set C = New cModule
   If Cnt = 1 And Mid(St, 1, 4) = "Type" Then
      ProjectType = Mid(St, 6, Len(St))
   End If
   If Mid(St, 1, 4) = "Name" Then
      ProjectName = Mid(St, 7, Len(St) - 7)
   End If
   

   If Mid(St, 1, 5) = "Class" Then
      C.ModuleType = "CLS"
      Ps = InStr(St, ";")
      C.ClassName = Mid(St, 7, Ps - 7)
      C.FileName = Mid(St, Ps + 2, Len(St))
   ElseIf Mid(St, 1, 6) = "Module" Then
      C.ModuleType = "BAS"
      Ps = InStr(St, ";")
      C.ClassName = Mid(St, 8, Ps - 8)
      C.FileName = Mid(St, Ps + 2, Len(St))
   ElseIf Mid(St, 1, 4) = "Form" Then
      C.ModuleType = "FRM"
      Ps = InStr(St, ";")
      If Ps > 0 Then
         C.ClassName = Mid(St, 6, Ps - 6)
         C.FileName = Mid(St, Ps + 2, Len(St))
      Else
         C.FileName = Mid(St, 6, Len(St))
         C.ClassName = C.FileName
      End If
   ElseIf Mid(St, 1, 11) = "UserControl" Then
      C.ModuleType = "CTL"
      Ps = InStr(St, ";")
      If Ps > 0 Then
         C.ClassName = Mid(St, 13, Ps - 13)
         C.FileName = Mid(St, Ps + 2, Len(St))
      Else
         C.FileName = Mid(St, 13, Len(St))
         C.ClassName = C.FileName
      End If
   End If
   If Len(C.ModuleType) > 0 Then
      C.Key = "R" & CStr(Files.Count + 1)
      Files.Add Item:=C, Key:=C.Key
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
   End Select
   
     
     Set itmx = TV.Nodes.Add(Root, tvwChild, Files(i).Key, Files(i).FileName)
     
     Set itmx = TV.Nodes.Add(Files(i).Key, tvwChild, "V" & Files(i).Key, "Declarations")
     Set itmx = TV.Nodes.Add("V" & Files(i).Key, tvwChild, "U1" & Files(i).Key, "User Types")
     Set itmx = TV.Nodes.Add("V" & Files(i).Key, tvwChild, "U2" & Files(i).Key, "Public")
     Set itmx = TV.Nodes.Add("V" & Files(i).Key, tvwChild, "U3" & Files(i).Key, "Private")
     Set itmx = TV.Nodes.Add(Files(i).Key, tvwChild, "P" & Files(i).Key, "Public Procedures")
     Set itmx = TV.Nodes.Add(Files(i).Key, tvwChild, "R" & Files(i).Key, "Private Procedures")
     
     Files(i).ResolveVariables
     
     For j = 1 To Files(i).UTypes.Count
        With Files(i).UTypes(j)
            Root$ = "U1" & Files(i).Key
            Key$ = Root$ & Format(j, "00")
            Set itmx = TV.Nodes.Add(Root$, tvwChild, Key, .Name)
            Root$ = Key
            For b = 1 To .Vars.Count
                Key$ = Root & Format(b, "00")
                Var$ = .Vars(b).Name & " As " & .Vars(b).varType
                Set itmx = TV.Nodes.Add(Root, tvwChild, Key, Var)
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
            Set itmx = TV.Nodes.Add(Root$, tvwChild, Key, Var)
        End With
     Next
     
     
     
     Files(i).ResolveProcedures
     For j = 1 To Files(i).Procs.Count
        With Files(i).Procs(j)
          If .IsPublic Then
              Root$ = "P" & Files(i).Key
          Else
              Root$ = "R" & Files(i).Key
          End If
            Key$ = Root$ & Format(j, "00")
            Set itmx = TV.Nodes.Add(Root$, tvwChild, Key, .FullName)
        End With
     Next
 Next

End Sub

Private Sub Command1_Click()
  With CommonDialog1
    .Filter = ".vbp files|*.vbp"
    .FileName = ""
    .ShowOpen
    Text1 = .FileName
    OpenVBP Text1
  End With
End Sub

Private Sub Command2_Click()
  Dim Indexfile As String
  Indexfile = App.Path & "\Index.html"
  Open Indexfile For Output As 1
  
  WriteHeader 1
  
'  Print #1, "<BODY BGCOLOR=""#FFFFFF"" TEXT#000000 style=""font-family: , Verdana"">"
'  Print #1, "<p align=""center""><strong><FONT SIZE=""5"" COLOR=""#000040"">" & ProjectName & "</FONT></strong></p>"
'  Print #1, "<HR SIZE=""5"" WIDTH=""100%"">"
  
  Print #1, "<BR><strong><FONT SIZE=""3"" COLOR=""#000040"">Forms</strong></FONT>"
  Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
  
  For i = 1 To Files.Count
      If Files(i).ModuleType = "FRM" Then
         Print #1, "<A HREF=""" & TrimExt(Files(i).ClassName) & ".htm"">" & TrimExt(Files(i).ClassName) & "</A>"
         Print #1, "<BR>"
      End If
  Next
  
  Print #1, "<BR><strong><FONT SIZE=""3"" COLOR=""#000040"">Object Classes</strong></FONT>"
  Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
  
  For i = 1 To Files.Count
      If Files(i).ModuleType = "CLS" Then
         Print #1, "<A HREF=""" & TrimExt(Files(i).ClassName) & ".htm"">" & TrimExt(Files(i).ClassName) & "</A>"
         Print #1, "<BR>"
      End If
  Next
  
  
  Print #1, "<BR><strong><FONT SIZE=""3"" COLOR=""#000040"">Modules</strong></FONT>"
  Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
  
  For i = 1 To Files.Count
      If Files(i).ModuleType = "BAS" Then
         Print #1, "<A HREF=""" & TrimExt(Files(i).ClassName) & ".htm"">" & TrimExt(Files(i).ClassName) & "</A>"
         Print #1, "<BR>"
      End If
  Next
  
  
  Print #1, "<BR>"
  Print #1, "<HR SIZE=""5"" WIDTH=""100%"">This page was created by the HTMLGen Compiler, made by Fabio Guerrazzi, <A HREF=""http://digilander.iol.it/WarZi/default.htm"">http://digilander.iol.it/WarZi/default.htm</A></BODY>"

  Print #1, "</HTML>"
  Close 1
  
  
' Save modules to html pages
  For i = 1 To Files.Count
      Files(i).ExportToHtml
  Next
  
  
  MsgBox "Html list done"
End Sub


Private Sub Command3_Click()
  With CommonDialog1
    .Filter = ".html files|*.html"
    .FileName = ""
    .ShowSave
    Text2 = .FileName
  End With

End Sub


Private Sub Form_Load()
 Text1 = ""
 Text2 = App.Path
End Sub


Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)
   Dim Msg As String
   Dim Ar() As String
   
   Caption = Node.Key
   
   Msg = Node & vbCrLf
   
   Msg = Msg & "Description: " & vbCrLf
   Msg = Msg & vbCrLf
   Msg = Msg & vbCrLf
   Msg = Msg & vbCrLf
   
   Msg = Msg & "Parameters: " & vbCrLf
   Msg = Msg & vbCrLf
   Msg = Msg & vbCrLf
   Msg = Msg & vbCrLf
   
   Msg = Msg & "Remarks: " & vbCrLf
   Msg = Msg & vbCrLf
   Msg = Msg & vbCrLf
   Msg = Msg & vbCrLf
   
   Msg = Msg & "See also: " & vbCrLf
   Msg = Msg & vbCrLf
   Msg = Msg & vbCrLf
   Msg = Msg & vbCrLf
   
   Msg = Msg & "---------------------"
   
  Text3 = Msg
End Sub
Function HexToDecimal(ByVal strHex As String) As Long

'This function is required by the function 'HTMLToRich'

'this function converts any hexidecimal color value
'(e.g. "0000FF" = Blue) to decimal color value.

Dim lngDecimal As Long, strCharHex As String, lngColor As Long
Dim lngChar As Long

If Left$(strHex$, 1) = "#" Then strHex$ = Right$(strHex$, 6)
  
strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)

  For lngChar& = Len(strHex$) To 1 Step -1
    strCharHex$ = Mid$(UCase$(strHex$), lngChar&, 1)
    
       Select Case strCharHex$
          Case 0 To 9
             lngDecimal& = CLng(strCharHex$)
          Case Else 'A,B,C,D,E,F
             lngDecimal& = CLng(Chr$((Asc(strCharHex$) - 17))) + 10
       End Select
       
    lngColor& = lngColor& + lngDecimal& * 16 ^ (Len(strHex$) - lngChar&)
  Next lngChar&
  
HexToDecimal = lngColor&

End Function




