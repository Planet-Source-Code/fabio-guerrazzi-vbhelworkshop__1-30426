VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public pType As String ' SUB,FUN,PRO
Public FullName As String
Public IsPublic As Boolean
Public Lines As New Collection
Public Vars As New Collection
Public Key As String

Public Caption As String
Public HelpDex As String
Public HelpRemarks As String
Public HelpTopic As String
Public HelpScreenShoot As String

Sub Save(hF As Long)

Dim fproc As rProc

 
 With fproc
  .Name = Name
  .pType = pType
  .FullName = FullName
  .IsPublic = IsPublic
  ReDim .Lines(Lines.Count)
  .nlines = Lines.Count
  .Key = Key
    .Caption = Caption
    .HelpDex = HelpDex
    .HelpRemarks = HelpRemarks
    .HelpTopic = HelpTopic
    .HelpScreenShoot = HelpScreenShoot
   For i = 1 To Lines.Count
      .Lines(i) = Lines(i)
   Next
 End With
 
 Put #hF, , fproc
 
End Sub

Sub OpenClass(hF As Long)

 Dim fproc As rProc, i&

 Get #hF, , fproc
  
 For i = 1 To Lines.Count
    Lines.Remove 1
 Next
  
 With fproc
  Name = .Name
  pType = .pType
  FullName = .FullName
  IsPublic = .IsPublic
  Key = .Key
  Caption = .Caption
  HelpDex = .HelpDex
  HelpRemarks = .HelpRemarks
  HelpTopic = .HelpTopic
  HelpScreenShoot = .HelpScreenShoot
  For i = 1 To .nlines
      Lines.Add Item:=.Lines(i)
  Next
 End With
 
 
End Sub
Private Sub Class_Initialize()
  Key = GenHandle
End Sub


