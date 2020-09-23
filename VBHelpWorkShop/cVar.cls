VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public varType As String
Public Mode As String  ' PUB,LOC
'Public vLine As String

Public Caption As String
Public HelpDex As String
Public HelpRemarks As String
Public HelpTopic As String
Public HelpScreenShoot As String

Sub Save(hF As Long)

Dim fVar As rVars

With fVar
  .Name = Name
  .varType = varType
  .Mode = Mode
      .Caption = Caption
    .HelpDex = HelpDex
    .HelpRemarks = HelpRemarks
    .HelpTopic = HelpTopic
    .HelpScreenShoot = HelpScreenShoot

End With

Put #hF, , fVar

End Sub

Sub OpenClass(hF As Long)

Dim fVar As rVars

Get #hF, , fVar

With fVar
  Name = .Name
  varType = .varType
  Mode = .Mode
  Caption = .Caption
  HelpDex = .HelpDex
  HelpRemarks = .HelpRemarks
  HelpTopic = .HelpTopic
  HelpScreenShoot = .HelpScreenShoot
End With


End Sub

