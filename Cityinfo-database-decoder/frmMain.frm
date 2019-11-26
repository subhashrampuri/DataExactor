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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   336
      Left            =   2700
      TabIndex        =   1
      Top             =   1908
      Width           =   1056
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   336
      Left            =   1656
      TabIndex        =   0
      Top             =   900
      Width           =   1128
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents objApp As Excel.Application
Attribute objApp.VB_VarHelpID = -1
Private WithEvents objBook As Excel.Workbook
Attribute objBook.VB_VarHelpID = -1
Private WithEvents objSheet As Excel.Worksheet
Attribute objSheet.VB_VarHelpID = -1
Private Sub Command1_Click()
Dim iRange As Integer
    Set objSheet = Nothing
    Set objBook = Nothing
    Set objApp = Nothing
    Set objApp = New Excel.Application
    Set objBook = objApp.Workbooks.Open(App.Path & "\MasterItinfo.xls", , True, , , , , , , False)
     
    Set objSheet = objBook.Worksheets("MasterItinfo")
     For iRange = 2 To 2535
        objSheet.Cells(iRange, 3) = fsDecode(objSheet.Cells(iRange, 2))
        objSheet.Cells(iRange, 5) = fsDecode(objSheet.Cells(iRange, 4))
        objSheet.Cells(iRange, 7) = fsDecode(objSheet.Cells(iRange, 6))
        objSheet.Cells(iRange, 14) = fsDecode(objSheet.Cells(iRange, 13))
        objSheet.Cells(iRange, 16) = fsDecode(objSheet.Cells(iRange, 15))
        objSheet.Cells(iRange, 22) = fsDecode(objSheet.Cells(iRange, 21))
        objSheet.Cells(iRange, 24) = fsDecode(objSheet.Cells(iRange, 23))
                
     Next

    objApp.Visible = xlSheetVisible
End Sub
Private Sub Command2_Click()
    'MsgBox Asc("ç")
    MsgBox Asc("$")
    MsgBox Chr(34)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set objSheet = Nothing
    Set objBook = Nothing
    Set objApp = Nothing
    
End Sub
Private Function fsDecode(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode = sTemp
End Function
Private Function fsReturnCharacter(ByVal iIndex As Integer) As String
    Dim vAlpha(80) As Variant
    vAlpha(0) = "a"
    vAlpha(1) = "b"
    vAlpha(2) = "c"
    vAlpha(3) = "d"
    vAlpha(4) = "e"
    vAlpha(5) = "f"
    vAlpha(6) = "g"
    vAlpha(7) = "h"
    vAlpha(8) = "i"
    vAlpha(9) = "j"
    vAlpha(10) = "k"
    vAlpha(11) = "l"
    vAlpha(12) = "m"
    vAlpha(13) = "n"
    vAlpha(14) = "o"
    vAlpha(15) = "p"
    vAlpha(16) = "q"
    vAlpha(17) = "r"
    vAlpha(18) = "s"
    vAlpha(19) = "t"
    vAlpha(20) = "u"
    vAlpha(21) = "v"
    vAlpha(22) = "w"
    vAlpha(23) = "x"
    vAlpha(24) = "y"
    vAlpha(25) = "z"
    vAlpha(26) = " "
    vAlpha(27) = "."
    vAlpha(28) = "A"
    vAlpha(29) = "B"
    vAlpha(30) = "C"
    vAlpha(31) = "D"
    vAlpha(32) = "E"
    vAlpha(33) = "F"
    vAlpha(34) = "G"
    vAlpha(35) = "H"
    vAlpha(36) = "I"
    vAlpha(37) = "J"
    vAlpha(38) = "K"
    vAlpha(39) = "L"
    vAlpha(40) = "M"
    vAlpha(41) = "N"
    vAlpha(42) = "O"
    vAlpha(43) = "P"
    vAlpha(44) = "Q"
    vAlpha(45) = "R"
    vAlpha(46) = "S"
    vAlpha(47) = "T"
    vAlpha(48) = "U"
    vAlpha(49) = "V"
    vAlpha(50) = "W"
    vAlpha(51) = "X"
    vAlpha(52) = "Y"
    vAlpha(53) = "Z"
    vAlpha(54) = "&"
    vAlpha(55) = ")"
    vAlpha(56) = "("
    vAlpha(57) = "-"
    vAlpha(58) = ","
    vAlpha(59) = "2"
    vAlpha(60) = "4"
    vAlpha(61) = "/"
    vAlpha(62) = "7"
    vAlpha(63) = "5"
    vAlpha(64) = "1"
    vAlpha(65) = "3"
    vAlpha(66) = "9"
    vAlpha(67) = "0"
    vAlpha(68) = ":"
    vAlpha(69) = "#"
    vAlpha(70) = "6"
    vAlpha(71) = "8"
    vAlpha(72) = "'"
    vAlpha(73) = "@"
    vAlpha(74) = "_"
    vAlpha(80) = "€"
    fsReturnCharacter = vAlpha(iIndex)
End Function
Private Function fiReturnIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(74) As Variant
    vAscii(0) = 162
    vAscii(1) = 178
    vAscii(2) = 177
    vAscii(3) = 165
    vAscii(4) = 179
    vAscii(5) = 167
    vAscii(6) = 163
    vAscii(7) = 164
    vAscii(8) = 166
    vAscii(9) = 94
    vAscii(10) = 175
    vAscii(11) = 176
    vAscii(12) = 138
    vAscii(13) = 188
    vAscii(14) = 189
    vAscii(15) = 128
    vAscii(16) = 127
    vAscii(17) = 134
    vAscii(18) = 190
    vAscii(19) = 126
    vAscii(20) = 131
    vAscii(21) = 135
    vAscii(22) = 129
    vAscii(23) = 136
    vAscii(24) = 143
    vAscii(25) = 144
    vAscii(26) = 159
    vAscii(27) = 153
    vAscii(28) = 182
    vAscii(29) = 235
    vAscii(30) = 181
    vAscii(31) = 237
    vAscii(32) = 245
    vAscii(33) = 239
    vAscii(34) = 254
    vAscii(35) = 236
    vAscii(36) = 169
    vAscii(37) = 243
    vAscii(38) = 170
    vAscii(39) = 241
    vAscii(40) = 171
    vAscii(41) = 238
    vAscii(42) = 191
    vAscii(43) = 195
    vAscii(44) = 240
    vAscii(45) = 212
    vAscii(46) = 242
    vAscii(47) = 158
    vAscii(48) = 244
    vAscii(49) = 141
    vAscii(50) = 246
    vAscii(51) = 140
    vAscii(52) = 90
    vAscii(53) = 89
    vAscii(54) = 197
    vAscii(55) = 248
    vAscii(56) = 247
    vAscii(57) = 226
    vAscii(58) = 225
    vAscii(59) = 61
    vAscii(60) = 228
    vAscii(61) = 214
    vAscii(62) = 231
    vAscii(63) = 229
    vAscii(64) = 63
    vAscii(65) = 227
    vAscii(66) = 233
    vAscii(67) = 64
    vAscii(68) = 210
    vAscii(69) = 196
    vAscii(70) = 230
    vAscii(71) = 232
    vAscii(72) = 250
    vAscii(73) = 211
    vAscii(74) = 36
    
    
        
    For iIndex = 0 To 74
        If iAscii = vAscii(iIndex) Then
             fiReturnIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnIndex = 80
End Function

