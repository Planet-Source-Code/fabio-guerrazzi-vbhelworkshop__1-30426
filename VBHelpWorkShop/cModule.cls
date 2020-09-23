VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Procs As New Collection
Public Vars As New Collection
Public UTypes As New Collection
Public Objects As New Collection ' Collezione di cobj che descrivono i Controlli di un form
Public Key As String

Public ModuleType As String
Public ClassName As String
Public FileName As String
  
Public Caption As String
Public HelpDex As String
Public HelpRemarks As String
Public HelpTopic As String
Public HelpScreenShoot As String


Sub CleanUP()
  Dim i As Long
  
  For i = 1 To Procs.Count
     Procs.Remove 1
  Next
  For i = 1 To Vars.Count
     Vars.Remove 1
  Next
  For i = 1 To Objects.Count
     Objects.Remove 1
  Next
  
End Sub

Sub ResolveObjects()
  Dim i As Long, St As String
  Dim C As cObj
  Dim Ar() As String
     
  If ModuleType <> "FRM" And ModuleType <> "CTL" Then Exit Sub
  
  Open FileName For Input As #1
  Do Until EOF(1)
     Line Input #1, St
     If Len(St) = 0 Then GoTo skip
     
     Ar = Split(Trim(St), " ")
     On Error Resume Next
     Count = -1
     Count = UBound(Ar)
     
     If Count >= 0 Then
        If Ar(0) = "Attribute" Then Exit Do ' Fine
        If Ar(0) = "End" Then GoTo skip
        If Count > 1 Then
        If Ar(0) = "Begin" And Ar(1) <> "VB.Form" And Ar(1) <> "VB.MDIForm" And Ar(1) <> "VB.Frame" Then
           Set C = New cObj
           C.Interactive = ParseControlType(Ar)
           C.ClassName = Ar(1)
           C.Name = Ar(2)
           C.GetChilds 1
           Objects.Add Item:=C
           Set C = Nothing
        End If
        End If
     End If
skip:
  Loop
  Close 1
End Sub


Sub ResolveProcedures()
  
  Dim St As String
  Dim Ar() As String, Count As Long
  Dim C As cProcedure, IProc As Long
  Dim Lin As Long
  Dim A2() As String
  Dim StartProc As Boolean
  
  Open FileName For Input As #1
  Do Until EOF(1)
     
     IProc = 0
     Lin = Lin + 1
     
     Line Input #1, St
     
     Ar = Split(Trim(St), " ")
     On Error Resume Next
     Count = -1
     Count = UBound(Ar)
     
     If Count >= 0 Then
        
        If ProcEnd(Ar) Then
           C.Lines.Add Item:=St
           Procs.Add Item:=C
           Set C = New cProcedure
        End If
        
        IProc = ProcStart(Ar)
        If IProc > 0 Then
           StartProc = True
           Set C = New cProcedure
           C.IsPublic = IsPublic(Ar)
           A2 = Split(Ar(IProc), "(")
           C.Name = A2(0)
           C.FullName = St
           C.Caption = C.FullName
        End If
        
        If StartProc Then C.Lines.Add Item:=St
        
        
     End If
  Loop
  
  Close 1

End Sub

Sub ResolveVariables()
  
  Dim St As String, i As Long
  Dim Ar() As String, Count As Long
  Dim C As cVars
  Dim U As UserType
  
  Open FileName For Input As #1
  Do Until EOF(1)
'    Set C = New cVars
     
     Line Input #1, St
     
     Ar = Split(Trim(St), " ")
     On Error Resume Next
     Count = -1
     Count = UBound(Ar)
     
     If Count >= 0 Then
        If ProcStart(Ar) > 0 Then Exit Do
        
        i = GetID(Ar(0))
        If InRange(i, 1, 6) Then
           If i = 6 Or GetID(Ar(1)) = 6 Then
              TypeOpen = True
              Set U = New UserType
              If i = 6 Then
                 U.Name = Ar(1)
              Else
                 U.Name = Ar(2)
              End If
           Else
              If InRange(i, 1, 2) Then ' Private
                 Set C = GetVar(Ar, 1)
                 If Not C Is Nothing Then
                    C.Mode = "Private"
                    Vars.Add Item:=C
                 End If
              Else ' Public
                 Set C = GetVar(Ar, 1)
                 If Not C Is Nothing Then
                    C.Mode = "Public"
                    Vars.Add Item:=C
                 End If
              End If
           End If
        ElseIf TypeOpen Then
           If Count >= 1 Then
            If Ar(0) = "End" And Ar(1) = "Type" Then
                UTypes.Add Item:=U
                TypeOpen = False
            End If
           End If
            If TypeOpen Then
              Set C = GetVar(Ar, 0)
              U.Vars.Add Item:=C
            End If
        End If
           
GoTo ex
        
        If TypeOpen Then
           If Ar(0) = "End" And Ar(1) = "Type" Then
              TypeOpen = False
           Else
              GoSub GetVar
           End If
        End If
        
        If Ar(0) = "Type" Then
           Set C = New cVars
        End If
        
        If Ar(0) = "Public" Or _
               Ar(0) = "Global" Then
           GoSub GetVar
           C.Mode = "PUB"
           Vars.Add Item:=C
        ElseIf Ar(0) = "Private" Or _
               Ar(0) = "Dim" Then
           GoSub GetVar
           C.Mode = "LOC"
           Vars.Add Item:=C
        End If
        
ex:
     End If
  Loop
  
  Close 1

Exit Sub

GetVar:
           If Ar(1) = "Enum" Or Ar(1) = "Const" Then
           ElseIf Ar(1) = "Type" Then
              C.Name = Ar(2)
              TypeOpen = True
           Else
              Set C = New cVars
              C.Name = Ar(1)
              If Count < 2 Then
                 C.varType = GetFixedType(C.Name)
                 If C.varType = 0 Then C.varType = "Variant"
              Else
                 C.varType = Ar(3)
              End If
           End If
Return
End Sub


Sub ExportToHtml()
  
  Dim HtmFile As String
  Dim St As String, Ps As Long, i As Long, j As Long
  Dim Ar() As String
  Dim ok As Boolean
  HtmFile = Prj.HTMLPath & "\" & TrimExt(ClassName) & ".htm"
  
  Open HtmFile For Output As 1
  
  
  WriteHeader 1
  
  
  Name$ = Caption 'TrimExt(ClassName)
  
  Print #1, "<p align=""center""><strong><FONT SIZE=""5"" COLOR=""#000040"">" & Name & "</FONT></strong></p>"
  Print #1, "<HR SIZE=""5"" WIDTH=""100%"">"
  
  If Len(HelpScreenShoot) > 0 Then WriteCenteredImage 1, HelpScreenShoot
  
  Print #1, "<BR><strong><FONT SIZE=""2.2"" COLOR=""#000040"">Members</strong></FONT>"
  Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
  
  For i = 1 To Procs.Count
      If PublicMBR Then ok = Procs(i).IsPublic Else ok = True
      If ok Then
         If SinglePage Then
             Lnk$ = Procs(i).Name & Format$(i, "00") & ".htm"
         Else
             Lnk$ = "#" & Procs(i).Name
         End If
         Print #1, "<a href=""" & Lnk & """>" & Procs(i).Name & "</a>"
         If SinglePage Then Print #1, "<BR>"
      End If
  Next
  
  Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
  
  Print #1, "<BR><strong><FONT SIZE=""2.2"" COLOR=""#000040"">Declarations</strong></FONT>"
  Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
  
  For i = 1 To Vars.Count
      'If PublicMBR Then Ok = Vars(i).IsPublic Else Ok = True
      ok = True
      Nome$ = Vars(i).Name & " As " & Vars(i).varType
      If ok Then Print #1, Nome & ", "
  Next
  
  Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
  
 If SinglePage Then
   PrintFooter 1
   Close 1
 End If
  
  
  For i = 1 To Procs.Count
   
      If PublicMBR Then ok = Procs(i).IsPublic Else ok = True
      If ok Then
      
       If SinglePage Then
            HtmFile = Prj.HTMLPath & "\" & Procs(i).Name & Format$(i, "00") & ".htm"
            Open HtmFile For Output As 1
            WriteHeader 1
            Name$ = Caption 'TrimExt(ClassName)
            Print #1, "<p align=""center""><strong><FONT SIZE=""5"" COLOR=""#000040"">" & Name & "." & Procs(i).Name & "</FONT></strong></p>"
            Print #1, "<HR SIZE=""5"" WIDTH=""100%"">"
       End If
         Print #1, "<BR>"
         Print #1, "<a name=""" & Procs(i).Name & """><font SIZE=""2.2"" COLOR=""#000000""><strong>" & Procs(i).FullName & "</strong></font></a>"
         
         j = ParseVars(Procs(i).FullName, Ar)
         If j > 0 Then
            WriteTableHead 1
            For k = 1 To j
                WriteTableRow 1, Ar(k - 1), "To Do"
            Next
            WriteTableClose 1
         End If
         If Len(Procs(i).HelpDex) > 0 Then Print #1, "<BR>&nbsp;&nbsp;&nbsp;" & Procs(i).HelpDex
         If Len(Procs(i).HelpRemarks) > 0 Then
            Print #1, "<BR>&nbsp;&nbsp;&nbsp;Remarks:"
            Print #1, "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Procs(i).HelpRemarks
         End If
         If Len(Procs(i).HelpScreenShoot) > 0 Then WriteCenteredImage 1, Procs(i).HelpScreenShoot
         Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
        
        If SinglePage Then
           PrintFooter 1
           Close 1
        End If
      End If
  Next
  
  
If Not SinglePage Then
 PrintFooter 1
  Close 1
End If
  

End Sub


Sub Save(hF As Long)
  Dim fMod As rModule
  With fMod
    .nProc = Procs.Count
    .nvars = Vars.Count
    .nObjects = Objects.Count
    .Key = Key
    .ModuleType = ModuleType
    .ClassName = ClassName
    .FileName = FileName
    .Caption = Caption
    .HelpDex = HelpDex
    .HelpRemarks = HelpRemarks
    .HelpTopic = HelpTopic
    .HelpScreenShoot = HelpScreenShoot
  End With
  
  Put #hF, , fMod
  
  For i = 1 To Procs.Count
       Procs(i).Save hF
  Next
  
  For i = 1 To Vars.Count
       Vars(i).Save hF
  Next
  
  For i = 1 To Objects.Count
       Objects(i).Save hF
  Next
  
End Sub

Sub OpenClass(hF As Long)
  
  Dim fMod As rModule
  
  Dim P As cProcedure
  Dim V As cVars
  Dim O As cObj
  
  CleanUP
  
  Get #hF, , fMod
  
  With fMod
    Key = .Key
    ModuleType = .ModuleType
    ClassName = .ClassName
    FileName = .FileName
    Caption = .Caption
    HelpDex = .HelpDex
    HelpRemarks = .HelpRemarks
    HelpTopic = .HelpTopic
    HelpScreenShoot = .HelpScreenShoot
      
    For i = 1 To .nProc
        Set P = New cProcedure
         P.OpenClass hF
         Procs.Add Item:=P, Key:=P.Key
        Set P = Nothing
    Next
    
    For i = 1 To .nvars
        Set V = New cVars
         V.OpenClass hF
         Vars.Add Item:=V ', Key:=V.Key
        Set V = Nothing
    Next
    
    For i = 1 To .nObjects
        Set O = New cObj
         O.OpenClass hF
         Objects.Add Item:=O, Key:=O.Key
        Set O = Nothing
    Next
    
  End With
  
End Sub


Sub WriteDoc()

' Scrive una pagina HTML con le definizioni dei controlli del form

  Dim HtmFile As String
  Dim St As String, Ps As Long, i As Long, j As Long
  Dim Ar() As String
  Dim ok As Boolean
  HtmFile = Prj.HTMLPath & "\" & TrimExt(ClassName) & ".htm"
  
  Open HtmFile For Output As 1
  
  WriteHeader 1
  
  Name$ = Caption 'TrimExt(ClassName)
  
 ' Print #1, "<BODY BGCOLOR=""#FFFFFF"" TEXT#000000 style=""font-family: , Verdana"">"
  Print #1, "<p align=""center""><strong><FONT SIZE=""5"" COLOR=""#000040"">" & Name & "</FONT></strong></p>"
  Print #1, "<HR SIZE=""5"" WIDTH=""100%"">"
  
  If Len(HelpScreenShoot) > 0 Then WriteCenteredImage 1, HelpScreenShoot
  If Len(HelpDex) > 0 Then
     Print #1, "<br>"
     Print #1, HelpDex
     Print #1, "<HR SIZE=""1"" WIDTH=""100%"">"
  End If
  
  ok = False
  For i = 1 To Objects.Count
    If Objects(i).ClassName = "VB.Menu" Then
       If Not ok Then
          WriteTableHead 1, "Menu"
          ok = True
       End If
         Objects(i).WriteRow 1, 1
    End If
  Next
  If ok Then WriteTableClose 1
  
  ok = False
  For i = 1 To Objects.Count
    If Objects(i).ClassName = "VB.TextBox" Then
       If Not ok Then
          WriteTableHead 1, "Input Fields"
          ok = True
       End If
       Objects(i).WriteRow 1, 1
    End If
  Next
  If ok Then WriteTableClose 1
  
  ok = False
  For i = 1 To Objects.Count
    If Objects(i).ClassName = "VB.CommandButton" Or Objects(i).ClassName = "MSComctlLib.Toolbar" Then
       If Not ok Then
          WriteTableHead 1, "Command Buttons"
          ok = True
       End If
       Objects(i).WriteRow 1, 1
    End If
  Next
  If ok Then WriteTableClose 1
  
  ok = False
  For i = 1 To Objects.Count
    If Objects(i).ClassName = "VB.OptionButton" Or Objects(i).ClassName = "VB.CheckBox" Then
       If Not ok Then
          WriteTableHead 1, "Options"
          ok = True
       End If
       Objects(i).WriteRow 1, 1
    End If
  Next
  If ok Then WriteTableClose 1
  
  ok = False
  For i = 1 To Objects.Count
    If Objects(i).ClassName = "VB.ListBox" Or Objects(i).ClassName = "MSComctlLib.TreeView" Or Objects(i).ClassName = "MSComctlLib.ListView" Or Objects(i).ClassName = "VB.ComboBox" Then
       If Not ok Then
          WriteTableHead 1, "Lists/Reports"
          ok = True
       End If
       Objects(i).WriteRow 1, 1
    End If
  Next
  If ok Then WriteTableClose 1
  

' **************************** Scrive tutti i controlli rimanenti
If Not Prj.SkipStaticControls Then
  ok = False
  For i = 1 To Objects.Count
    If Objects(i).Done = 0 Then
       If Not ok Then
          WriteTableHead 1, "Others"
          ok = True
       End If
       Objects(i).WriteRow 1, 1
    End If
  Next
  If ok Then WriteTableClose 1
End If

' **************************** END
  
PrintFooter 1
Close 1

End Sub


Private Sub Class_Initialize()
  Key = GenHandle
End Sub


