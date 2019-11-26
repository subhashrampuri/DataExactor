VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCount 
      Height          =   288
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1212
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count"
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1092
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1212
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Private Sub cmdCount_Click()
   MsgBox i
End Sub

Private Sub cmdStart_Click()
 On Error GoTo Localerr
 
    Dim objCon As New ADODB.Connection
    Dim objRs As New ADODB.Recordset
    Dim objApp As New Excel.Application
    Dim objBook As Excel.Workbook
    Dim objSheet As Excel.Worksheet
    
    
    objCon.Open "DSN=test;UID=sa;PWD=;"
    
    i = 0
    
    Set objBook = objApp.Workbooks.Open(App.Path & "\test.xls")
    
    Set objSheet = objBook.Sheets(1)
    
    For i = 2 To CLng(txtCount.Text)

        If Len(Trim(CStr(objSheet.Cells(i, 1)))) <> 0 Then
            objRs.Open "SELECT com_sDesignation from tblDesignation where com_sName='" & CStr(objSheet.Cells(i, 1)) & "'", objCon, adOpenStatic, adLockReadOnly, adCmdText
        
             If objRs.EOF = False Then
                objSheet.Cells(i, 5) = objRs(0)
             End If
    
            If objRs.State = 1 Then
                objRs.Close
            End If
        End If
    Next
    
    objBook.Save
    
    MsgBox "Done"
    
    Set objSheet = Nothing
    Set objBook = Nothing
    objApp.Quit
    Set objApp = Nothing
    
    Set objRs = Nothing
    Set objCon = Nothing
    Exit Sub
Localerr:
    'MsgBox Err.Description & "i=" & i
    
     If objRs.State = 1 Then
        objRs.Close
     End If
    Resume Next
End Sub
