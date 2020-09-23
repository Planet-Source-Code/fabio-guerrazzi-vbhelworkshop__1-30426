VERSION 5.00
Begin VB.Form fCat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Categories"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "&Remove"
      Height          =   315
      Left            =   1740
      TabIndex        =   3
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Update"
      Height          =   315
      Left            =   780
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   675
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   3435
   End
End
Attribute VB_Name = "fCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub FillList()

  List1.Clear
  For i = 1 To Prj.nCat
     List1.AddItem Prj.Cats(i)
  Next
End Sub

Private Sub Command1_Click()
 Dim St As String
 St = InputBox("New Category")
 If Len(St) = 0 Then Exit Sub
 Prj.nCat = Prj.nCat + 1
 ReDim Preserve Prj.Cats(nCat)
 Prj.Cats(nCat) = St
 FillList
End Sub


Private Sub Command2_Click()
 Dim St As String
 i = List1.ListIndex
 If i = -1 Then Exit Sub
 
 St = List1.List(i)
 St = InputBox("Update Category", , St)
 If Len(St) = 0 Then Exit Sub
 Prj.Cats(i + 1) = St
 FillList

End Sub

Private Sub Command3_Click()
 Dim St As String
 i = List1.ListIndex
 If i = -1 Then Exit Sub
 St = List1.List(i)
 If MsgBox("Remove " & St & " ?", 32 + 4, "Delete item") = vbYes Then
    For j = i + 2 To Prj.nCat - 1
        Prj.Cats(j) = Prj.Cats(j + 1)
    Next
    Prj.nCat = Prj.nCat - 1
    ReDim Preserve Prj.Cats(Prj.nCat)
    FillList
 End If

End Sub


Private Sub Form_Load()
  FillList
End Sub


