VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7008
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9456
   LinkTopic       =   "Form3"
   ScaleHeight     =   7008
   ScaleWidth      =   9456
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   6576
      Left            =   3288
      TabIndex        =   1
      Top             =   72
      Width           =   5892
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   444
      Left            =   96
      TabIndex        =   0
      Top             =   144
      Width           =   2988
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Dim i As Integer
    
    For i = 0 To 255
        List1.AddItem "i=" & i & " " & Chr(i)
    Next
End Sub
