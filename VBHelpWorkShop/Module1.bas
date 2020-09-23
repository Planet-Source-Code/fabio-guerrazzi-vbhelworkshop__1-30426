Attribute VB_Name = "Module1"
Public LastFile As String
Public LinkedText As String
Public LinkedURL As String
Public Link1 As Variant

Public SinglePage As Boolean

Type rParm
 ProjectType As String ' Tipo di progetto
 ProjectName As String
 ProjectTitle As String
 ProjectFile As String
 ProjectVersion As String
 HTMLPath As String
 PathVBP As String
 PublicMBR As Boolean
 OutputMode As Long ' 0-Progr. Reference, 1-App Documentation
 BackGroundImage As String
 SkipStaticControls As Boolean
 IncludeEmptyItems As Boolean
 GenHHP As Boolean
 AddFiles(5) As String ' copyright, whatsnew page etc..
 AdFlag(5) As Boolean
 Cats() As String
 nCat As Long
 ModulesCount As Long
 Author As String
 URL As String
End Type

Public Prj As rParm

Public Files As New Collection

Public License As String
Public Demo As Boolean

Public CurObj As Object

Type rVars
  Name As String
  varType As String
  Mode As String  ' PUB,LOC
  Caption As String
  HelpDex As String
  HelpRemarks As String
  HelpTopic As String
  HelpScreenShoot As String
End Type

Type rProc
  Name As String
  pType As String ' SUB,FUN,PRO
  FullName As String
  IsPublic As Boolean
  Lines() As String
  nlines As Long
  Key As String

  Caption As String
  HelpDex As String
  HelpRemarks As String
  HelpTopic As String
  HelpScreenShoot As String

End Type

Type rObjects
  ClassName As String
  Name As String
  nChilds As Long
  Caption As String
  Done As Long
  Key As String
  Interactive As Boolean
  HelpDex As String
  HelpRemarks As String
  HelpTopic As String
  HelpScreenShoot As String
End Type

Type rModule
  nProc As Long
  nvars As Long
  nObjects As Long
  Key As String
  ModuleType As String
  ClassName As String
  FileName As String
  
  Caption As String
  HelpDex As String
  HelpRemarks As String
  HelpTopic As String
  HelpScreenShoot As String

End Type

Function CatchImage(File As String) As String
  
  ' Verifica ed eventualmente copia l'immagine
  ' nella cartella di destinazione
  ' Retituisce il solo nomefile (es: miofile.jpg)
  
  Dim Img As String
  Dim NewPos As String
  
  If Len(File) > 0 Then
     Img = ExtractFileFromString(File)
     NewPos = Prj.HTMLPath & "\" & Img
     If Not FileExists(NewPos) Then FileCopy File, NewPos
     CatchImage = Img
  End If

End Function

Function ExtractFileFromString(Fn As String) As String
    Dim Ps As Byte
    Dim Bps As Byte
    
    Ps = 0
    Do
     Bps = Ps
     Ps = InStr(Ps + 1, Fn, "\")
    Loop While Ps > 0
    
    If Bps > 0 Then
       ExtractFileFromString = Mid$(Fn, Bps + 1, Len(Fn))
    Else
       ExtractFileFromString = Fn
    End If

End Function




Function GenHandle() As String
  Randomize
  
  GenHandle = "H" & Hex((Rnd * 256) + 1) & Hex((Rnd * 256) + 1) & Hex((Rnd * 256) + 1) & Hex((Rnd * 256) + 1)

End Function


Function GetIcon(ClassName As String) As Long
 ' Icone:
 ' 9= Menu
 ' 13= Buttons
 ' 14=Text
 ' 15=lists
 ' 16 = options

 Select Case ClassName
    Case "VB.TextBox": GetIcon = 14
    Case "VB.Menu": GetIcon = 8
    Case "VB.CommandButton":  GetIcon = 13
    Case "MSComctlLib.Toolbar": GetIcon = 28
    Case "VB.OptionButton", "VB.CheckBox"
      GetIcon = 16
    Case "VB.Label":   GetIcon = 33
    Case "VB.ListBox":   GetIcon = 22
    Case "MSComctlLib.TreeView": GetIcon = 29
    Case "MSComctlLib.ListView": GetIcon = 23
    Case "VB.ComboBox": GetIcon = 21
    Case "VB.PictureBox", "VB.Image": GetIcon = 20
    Case "VB.VScrollBar", "VB.HScrollBar": GetIcon = 18
    Case "MSComDlg.CommonDialog": GetIcon = 32
    Case "MSComctlLib.ListImages": GetIcon = 25
    Case "MSComctlLib.Slider": GetIcon = 26
    Case "MSComctlLib.ProgressBar": GetIcon = 24
    Case Else
      GetIcon = 19
 End Select
End Function

Sub GetNames(St As String, CN As String, Fn As String)

Dim Ps As Long, CS As String

CN = ""
Fn = ""

Ps = InStr(St, "=")
If Ps = 0 Then Exit Sub
CS = Mid(St, Ps + 1, Len(St))
Ps = InStr(CS, ";")
If Ps = 0 Then
   Fn = CS
   CN = EstraiNomeFile(ExtractFileFromString(CS))
Else
   CN = Mid(CS, 1, Ps - 1)
   Fn = Mid(CS, Ps + 1, Len(CS))
End If
Fn = Trim(Fn)
CN = Trim(CN)
'If Mid(Fn, 1, 2) = ".." Then Fn = Replace(Fn, "..", PathVBP)


End Sub


Function EstraiNomeFile(St As String) As String
 ' Elimina l'estensione ad un nome di file
  
  Dim St1 As String
  Dim Ps As Integer
  
  St1 = ExtractFileFromString(St) ' Toglie la Path
  
  Ps = InStr(St1, ".")
  If Ps > 0 Then
     EstraiNomeFile = Mid$(St1, 1, Ps - 1)
  Else
     EstraiNomeFile = St1
  End If

End Function


Function GetFixedType(V As String, Optional Trail = 0) As String
  t$ = Mid(V, Len(V), 1)
  Select Case t
     Case "%": GetFixedType = "Integer"
     Case "!": GetFixedType = "Single"
     Case "#": GetFixedType = "Double"
     Case "&": GetFixedType = "Long"
  End Select
End Function


Function ExtractPathFromString(St As String) As String

' Restituisce la sola path contenuta in una stringa

    Dim Ps As Byte
    Dim Bps As Byte
    
    Ps = 0
    Do
     Bps = Ps
     Ps = InStr(Ps + 1, St, "\")
    Loop While Ps > 0
    
    If Bps > 0 Then
       ExtractPathFromString = Mid$(St, 1, Bps - 1) ', Len(St))
    Else
       ExtractPathFromString = ""
    End If


End Function


Function GetID(Code As String) As Long

  Select Case UCase(Code)
    Case "PRIVATE": GetID = 1
    Case "DIM": GetID = 2
    Case "DECLARE": GetID = 3
    Case "PUBLIC": GetID = 4
    Case "GLOBAL": GetID = 5
    Case "TYPE": GetID = 6
    
    Case "SUB": GetID = 10
    Case "FUNCTION": GetID = 11
    Case "PROPERTY": GetID = 12
    Case "IF": GetID = 20
    Case "THEN": GetID = 21
    Case "ELSE": GetID = 22
    Case "DO": GetID = 23
    Case "WHILE": GetID = 24
    Case "UNTIL": GetID = 25
    Case "LOOP": GetID = 26
    Case "WITH": GetID = 27
    Case "END": GetID = 2000
    Case "'": GetID = 2001
    Case "EXIT": GetID = 2002
  
  End Select

End Function


Function GetVar(Ar, ist As Long) As cVars
    
    Dim C As New cVars
           If Ar(ist) = "Enum" Or Ar(ist) = "Const" Then
              Set GetVar = Nothing
              Exit Function
           Else
              Set C = New cVars
              C.Name = Ar(ist)
              If UBound(Ar) < 2 Then
                 C.varType = GetFixedType(C.Name)
                 If C.varType = 0 Then C.varType = "Variant"
              Else
                 C.varType = Ar(ist + 2)
                 If C.varType = "New" Then C.varType = Ar(ist + 3)
              End If
           End If

    Set GetVar = C
    
End Function


Function InRange(i As Long, i1 As Long, i2 As Long) As Boolean
  InRange = i >= i1 And i <= i2
End Function

Function ParseControlType(Ar() As String) As Boolean
  C& = UBound(Ar)
  
  If Ar(0) <> "Begin" Then Exit Function
  
  If Ar(1) = "VB.TextBox" Or _
     Ar(1) = "VB.OptionButton" Or _
     Ar(1) = "VB.Menu" Or _
     Ar(1) = "VB.CommandButton" Or _
     Ar(1) = "VB.ComboBox" Or _
     Ar(1) = "MSComctlLib.TreeView" Or _
     Ar(1) = "MSComctlLib.ListView" Or _
     Ar(1) = "VB.ListBox" Then
     ParseControlType = True
  End If

  
End Function

Function ParseVars(St As String, Ar() As String) As Long
 ' Estrae le variabili interne ad una Sub, Function, property ecc.
 ' NB. per Ubound(Ar)=0 la funzione restituisce 1
    p1& = InStr(St, "(")
    p2& = InStr(St, ")")
    l& = (p2 - p1) - 1
    Erase Ar
    If l > 0 Then
       ss$ = Mid(St, p1 + 1, l)
       If InStr(ss, ",") Then
          Ar = Split(ss, ",")
          ParseVars = UBound(Ar) + 1
       Else
          ReDim Ar(0)
          Ar(0) = ss
          ParseVars = 1
       End If
    End If
   
End Function

Sub PrintFooter(hF As Long)
  
  Print #hF, "<BR>"
  Print #hF, "<HR SIZE=""5"" WIDTH=""100%"">"
'  URL$ = "Created by VB Help WorkShop 1.0 <A HREF=""http://www.mandix.com"">http://www.mandix.com</A> © 2001"
  URL$ = "Copyright ©" & Year(Now) & " <A HREF=""" & Prj.URL & """>" & Prj.Author & "</A> Help created on " & Format(Now)
  Print #hF, "<p align=""center""><FONT SIZE=""1"" COLOR=""#808080""><small><small>" & URL & "</FONT></strong></small></p>"
  Print #hF, "</BODY>"
  
  Print #hF, "</small>"
  Print #hF, "</HTML>"
  
End Sub

Function ProcStart(Ar() As String) As Long
   Dim i1 As Long, i As Long
   On Error Resume Next
   '       10-12= Sub,Function etc
   
   For i = 0 To UBound(Ar)
       i1 = GetID(Ar(i))
       If i1 = 2000 Or i1 = 2001 Or i1 = 2002 Then Exit Function
       If InRange(i1, 10, 11) Then
          ProcStart = i + 1
          Exit Function
       ElseIf i1 = 12 Then
          ProcStart = i + 2
          Exit Function
       End If
   Next
   
'   i1 = GetID(Ar(0))
'   i2 = GetID(Ar(1))
'
'   If InRange(i1, 1, 5) And InRange(i2, 10, 12) Then
'      ProcStart = 2
'   ElseIf InRange(i1, 10, 12) Then
'      ProcStart = 2
'   End If

End Function


Function IsPublic(Ar() As String) As Long
   Dim i1 As Long, i As Long
   On Error Resume Next
   
   IsPublic = True
   For i = 0 To UBound(Ar)
       i1 = GetID(Ar(i))
       If i1 = 2000 Or i1 = 2001 Then Exit Function
       If InRange(i1, 4, 5) Then
          Exit Function
       ElseIf InRange(i1, 1, 2) Then
          IsPublic = False
          Exit Function
       End If
   Next

End Function
Function ProcEnd(Ar() As String) As Boolean
   
   Dim i1 As Long, i2 As Long
   On Error Resume Next
   ' range 1-5= Dichiaratori
   '       10-12= Sub,Function etc
   i1 = GetID(Ar(0))
   i2 = GetID(Ar(1))
   
   ProcEnd = i1 = 2000 And InRange(i2, 10, 12)

End Function
Sub SavePrj(FileName As String)
  Open FileName For Binary As 1
  
  Prj.ModulesCount = Files.Count
  
  Put 1#, , Prj
    
    For i = 1 To Files.Count
        Files(i).Save 1
    Next
  
  Close 1
End Sub

Sub OpenPrj(FileName As String)
  
  Dim C As cModule
  
  Open FileName For Binary As 1
  
  Get 1#, , Prj
  
  For i = 1 To Prj.ModulesCount
      Set C = New cModule
      C.OpenClass 1
      Files.Add Item:=C, Key:=C.Key
      Set C = Nothing
  Next
  
  Close 1
End Sub

Function TrailAP(St As String) As String
       TrailAP = Replace(St, Chr(34), "")
End Function


Function TrimExt(V As String) As String

  Dim Ps As Long
  Ps = InStr(V, ".")
  If Ps = 0 Then
    TrimExt = V
  Else
    TrimExt = Mid$(V, 1, Ps - 1)
  End If

End Function


Sub WriteCenteredImage(hF As Long, File As String)
Dim Img As String

Img = CatchImage(File)
If Len(Img) > 0 Then Print #hF, "<p align=""center""><img src=""" & Img & """></p>"

End Sub

Sub ChangeFont(f As Object)
  Dim x As Object
  
  On Error Resume Next
  
  For Each x In f
      If x Is Nothing Then Exit For
      If Len(x.Name) > 0 Then
        If x.Font.Name = "MS Sans Serif" Then x.Font.Name = "Tahoma"
        If Not x Is Nothing Then ChangeFont x
      End If
  Next

End Sub


Sub WriteHeader(hFile As Long)

  Dim ImgStr As String, Img As String
  
  Name$ = Prj.ProjectName
  If Len(Prj.BackGroundImage) > 0 Then
     Img = CatchImage(Prj.BackGroundImage)
     ImgStr = "background=""" & Img & """"
  End If
  
  Print #hFile, "<HTML>"
  Print #hFile, "<small>"
'  Print #hFile, "<body bgcolor=""#FFFFFF"" " & ImgStr & ">"
  Print #hFile, "<body bgcolor=""#FFFFFF"" style=""font-family: , Verdana"" " & ImgStr & ">"
  Print #hFile, "<p style=""background-color: rgb(107,183,211)""><em><font SIZE=""4"" COLOR=""#000040"">" & Name & "</font></em></p>"
  
End Sub


Function FileExists(FileName As String) As Integer
 
Dim i As Integer
On Error Resume Next
If Len(FileName) = 0 Then Exit Function

i = Len(Dir$(FileName))
If Err Or i = 0 Then
    FileExists = False
  Else
    FileExists = True
End If

End Function



Sub WriteTableClose(hF As Long)
 Print #hF, "</table>"
End Sub

Sub WriteTableHead(hF As Long, Optional ColHead1 = "Parameter")
'Print #hF, " <p><small>Parameters:</small></p>"
 
 Head1$ = ColHead1
 Head2$ = "Description"

Print #hF, "<BR>"
Print #hF, "<table border=""0"" width=""100%"" style=""border: 1px"">"
Print #hF, "  <tr>"
Print #hF, "    <th width=""20%"" bgcolor=""#D8E7E9"""
Print #hF, "    Style = ""background-color: rgb(147,158,240); color: rgb(255,255,255); border: 1px solid""> <small>" & Head1 & "</small></th>"
Print #hF, "    <th width=""80%"""
Print #hF, "    Style = ""background-color: rgb(147,158,240); color: rgb(255,255,255); border: 1px solid""> <small>" & Head2 & "</small></th>"
Print #hF, "   </tr>"

End Sub


Sub WriteTableRow(hF As Long, Par As String, Dex As String, Optional Img As String = "")
Dim DxImg As String
If Len(Img) > 0 Then DxImg = "<img src = """ & Img & """>"
Print #hF, "   <tr>"
Print #hF, "    <td width=""20%"" bgcolor=""#D8E7E9"" style=""border: 1px solid""><small>" & Par & "</small></td>"
Print #hF, "    <td width=""80%"" style=""border: 1px solid""><small>" & Dex & "</small>" & DxImg & "</td>"
'If Len(Img) > 0 Then WriteCenteredImage hF, Img
Print #hF, "      </tr>"
'  <tr>
'    <td width="20%" bgcolor="#D8E7E9" style="border: 1px solid"></td>
'    <td width="80%" style="border: 1px solid"></td>
'  </tr>

End Sub


