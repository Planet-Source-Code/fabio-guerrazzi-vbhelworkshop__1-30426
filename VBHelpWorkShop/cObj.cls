VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ClassName As String
Public Name As String
Public Childs As New Collection
Public Caption As String
Public Done As Long
Public Key As String
Public Interactive As Boolean

Public HelpDex As String
Public HelpRemarks As String
Public HelpTopic As String
Public HelpScreenShoot As String

Sub AddToTree(TV As Object, Root As String)
 
 
 
 Dim itmx As Node, Dex As String
 Dim ClassName2 As String
 Dim ok As Boolean
 Dim Ico As Long
 
 Ico = GetIcon(ClassName)
 
 On Error Resume Next
 
 
 Ps& = InStr(ClassName, ".")
 If Ps > 0 Then
    ClassName2 = Mid(ClassName, Ps + 1, 100)
 Else
    ClassName2 = ClassName
 End If
  
 ok = Prj.SkipStaticControls And Interactive Or Not Prj.SkipStaticControls
 If ok Then
    Set itmx = TV.Nodes.Add(Root, tvwChild, Key, ClassName2 & ": " & Name & "(" & Caption & ")", Ico, Ico)
    Set itmx = Nothing
    For i = 1 To Childs.Count
       Childs(i).AddToTree TV, Key
    Next
 End If
 
 
End Sub

Sub GetChilds(hF As Long)
    Dim C As cObj
    Dim St As String
    Dim Ar() As String
    
    Key = GenHandle
    
    Do Until EOF(hF)
     
     Line Input #hF, St
     Ar = Split(Trim(St), " ")
     On Error Resume Next
     Count = -1
     Count = UBound(Ar)
     If Count >= 0 Then
        If Ar(0) = "End" Then
           If Len(Caption) > 0 Then Caption = TrailAP(Caption) Else Caption = Name
           Exit Sub
        End If
        If Ar(0) = "Caption" Then
           Ps = InStr(St, "=")
           If Ps > 0 Then Caption = Trim(Mid(St, Ps + 1, Len(St)))
        End If
        If Count > 1 Then
        If Ar(0) = "Begin" And Ar(1) <> "VB.Form" And Ar(1) <> "VB.MDIForm" And Ar(1) <> "VB.Frame" Then
          ' If Len(Ar(2)) = 0 Then Stop
           Set C = New cObj
           C.Interactive = ParseControlType(Ar) ' True se è un controllo di input o di selezione
           C.ClassName = Ar(1)
           C.Name = Ar(2)
           C.GetChilds hF
           Childs.Add Item:=C
           Set C = Nothing
        End If
        End If
     End If
    
  Loop
  
  If Len(Caption) > 0 Then Caption = TrailAP(Caption) Else Caption = Name
  
End Sub

Function GetObj(SearchKey As String) As cObj
  If Key = SearchKey Then
     Set GetObj = Me
  Else
     For i& = 1 To Childs.Count
         Set GetObj = Childs(i).GetObj(SearchKey)
         If Not GetObj Is Nothing Then Exit Function
     Next
  End If
End Function

Sub Save(hF As Long)

Dim fobj As rObjects, i&

With fobj
  .ClassName = ClassName
  .Name = Name
  .nChilds = Childs.Count
  .Caption = Caption
  .Done = Done
  .Key = Key
  .Interactive = Interactive
  
  .HelpDex = HelpDex
  .HelpRemarks = HelpRemarks
  .HelpTopic = HelpTopic
  .HelpScreenShoot = HelpScreenShoot
  
End With

Put #hF, , fobj

For i = 1 To Childs.Count
   Childs(i).Save hF
Next


End Sub

Sub OpenClass(hF As Long)

Dim fobj As rObjects, i&, C As cObj

Get #hF, , fobj

For i = 1 To Childs.Count
  Childs.Remove 1
Next

With fobj
  ClassName = .ClassName
  Name = .Name
 ' nChilds = Childs.Count
  Caption = .Caption
  'Done = Done
  Key = .Key
  Interactive = .Interactive
  
  HelpDex = .HelpDex
  HelpRemarks = .HelpRemarks
  HelpTopic = .HelpTopic
  HelpScreenShoot = .HelpScreenShoot
  
End With


For i = 1 To fobj.nChilds
  Set C = New cObj
   C.OpenClass hF
   Childs.Add Item:=C, Key:=C.Key
  Set C = Nothing
Next


End Sub

Sub WriteHelp(hF As Long)
   Dim i As Long, Dex As String
   
   
   
   
   
   
   
   
   
 '  Print #hF, "<BR>"
   Print #hF, "<a name=""" & Name & """><font SIZE=""1"" COLOR=""#000040""><strong>" & Caption & "</strong></font></a>"
  ' Print #hF, "<BR><font color=""#008000""><small>&nbsp;&nbsp;&nbsp;Usage <br></font></small>"
   Print #hF, "<BR><small>" & HelpDex & "</small>"
   
   If Len(HelpRemarks) > 0 Then
       Print #hF, "<BR><font color=""#008000""><small>&nbsp;&nbsp;&nbsp;Remarks: <br></font></small>"
       Print #hF, "<small>" & HelpRemarks & "</small>"
    '   Print #hF, HelpRemarks
   End If

   If Len(HelpScreenShoot) > 0 Then WriteCenteredImage hF, HelpScreenShoot
   
   Print #hF, "<HR SIZE=""1"" WIDTH=""100%"">"

  For i = 1 To Childs.Count
      Childs(i).WriteHelp hF
  Next

   Print #hF, "<BR>"

End Sub



Sub WriteHTMLHeader(hF As Long, Ident As Long)
  Dim Tabs As String, i As Long, j As Long
  Dim Dex As String
  
  Done = 1 ' Informa il programma che ha già eseguito la scrittura di questo oggetto
  
  Const TB = "&nbsp;&nbsp;&nbsp;"
  
  For i = 1 To Ident
     Tabs = Tabs & TB
  Next
  
 ' If Len(Caption) > 0 Then Dex = TrailAP(Caption) Else Dex = Name
  
  Print #hF, "<BR>" & Tabs & "<a href=""#" & Name & """>" & Caption & "</a>"
  
  j = Ident + 1
  For i = 1 To Childs.Count
      Childs(i).WriteHTMLHeader hF, j
  Next

End Sub


Sub WriteRow(hF As Long, Ident As Long)
  Dim Tabs As String, i As Long, j As Long
  Dim Dex As String
  
  If Len(Caption) = 0 Then Exit Sub
  
  Const TB = "&nbsp;&nbsp;&nbsp;"
  
  Done = 1
  
  For i = 1 To Ident - 1
     Tabs = Tabs & TB
  Next

   If Len(HelpDex) > 0 Then Dex = HelpDex Else Dex = "To Do"
   
   WriteTableRow hF, Tabs & Caption, Dex, HelpScreenShoot
   
  j = Ident + 1
  For i = 1 To Childs.Count
      Childs(i).WriteRow hF, j
  Next

End Sub


