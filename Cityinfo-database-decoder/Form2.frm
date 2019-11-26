VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form2"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   636
      Left            =   504
      TabIndex        =   1
      Top             =   1584
      Width           =   2052
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   348
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1380
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
 MsgBox Asc("†")
End Sub
Private Sub Command2_Click()
    
    'MsgBox fsDecode_City("¬«£ïπ∫∫≈∂¿ï¿ƒ ¡")
    'MsgBox fsDecode_State("¬«£ïπ∫∫≈∂¿ï¿ƒ ¡")
     'MsgBox fsDecode_ContactPerson("»∂√ø∂Œï¬∂¡Ωƒ…«∂°ïπæ«∫∏…ƒ«")
     MsgBox fsDecode_Company("ø∂ºπæ»Ωï»æ√ºΩïºæ¡¡°ïπæ«∫∏…ƒ«")
     ' MsgBox fsDecode_Profiles("∆∫«Œø∫ºÕŒæÀÀô»øôÃ“«Õ¡æÕ¬º®Ã…æº¬∫≈¬Õ“ô≈ŒªÀ¬º∫«ÕÃôüô¿Àæ∫ÃæÃôª∫ÃæΩô»«ôÃ¬≈¬º»«æ•ô∆»≈“ÃŒ≈…¡¬Ωæ•ôÃ“«Õ¡æÕ¬ºô»¬≈ÃôæÕºßôø»Àô∆∫º¡¬«æÃ•ôæ Œ¬…∆æ«Õôüôº»∆…»«æ«ÕÃô–»Àƒ¬«¿ôŒ«ΩæÀôÃæœæÀæôº»«Ω¬Õ¬»«Ãô»øô¡¬¿¡ôÕæ∆…æÀ∫ÕŒÀæ•ô≈»∫Ω•ô∆»¬ÃÕŒÀæ•ôΩŒÃÕôæÕºßô¬«ΩŒÃÕÀ¬∫≈ô∆∫¬«Õæ«∫«ºæô∫æÀ»Ã")
     'MsgBox fsDecode_Address1("º®≤´ßõ√º–’õ∆√ºŒß")
     'MsgBox fsDecode_Address2("øæøæœùœÃæ¡")
     'MsgBox fsDecode_Address3(" »—”»üÕ¿∆¿—ü»Õ√‘“”—»¿Àü¿—ƒ¿")
     'MsgBox fsDecode_Pincode("∑∏∏ßπºªß∫∫∫æº")
     'MsgBox fsDecode_Phone("∑∏∏¥πΩºΩ∑∏∑∏∂πΩºΩ∑∏∑π")
     'MsgBox fsDecode_Fax("π∫∫ªªæΩΩπ¿º")
     'MsgBox fsDecode_Website("Ûˇˇ˚≈∫∫πˇÛÏˆ˝æπÙˇ˘Ô˝˛πÓ˙¯")
     MsgBox fsDecode_Email("ÊÿÁ€‹°Ê°Â≥Ï‘€‚‚°÷‚°‹·")
End Sub
Private Function fsDecode_City(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnCityIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_City = sTemp
End Function
Private Function fsDecode_State(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnStateIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_State = sTemp
End Function
Private Function fsDecode_ContactPerson(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnContactPersonIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_ContactPerson = sTemp
End Function
Private Function fsDecode_Company(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnCompanyIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_Company = sTemp
End Function
Private Function fsDecode_Profiles(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnProfilesIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_Profiles = sTemp
End Function
Private Function fsDecode_Address1(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnAddress1Index(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_Address1 = sTemp
End Function
Private Function fsDecode_Address2(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnAddress2Index(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_Address2 = sTemp
End Function
Private Function fsDecode_Address3(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnAddress3Index(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_Address3 = sTemp
End Function
Private Function fsDecode_Pincode(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnPinCodeIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_Pincode = sTemp
End Function
Private Function fsDecode_Phone(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnPhoneIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_Phone = sTemp
End Function
Private Function fsDecode_Fax(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnFaxIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_Fax = sTemp
End Function
Private Function fsDecode_Website(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnWebsiteIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_Website = sTemp
End Function
Private Function fsDecode_Email(ByVal sText As String) As String
    Dim vArray
    Dim sTemp As String
    Dim i As Integer
    For i = 1 To Len(sText)
        sTemp = sTemp & fsReturnCharacter(fiReturnEmailIndex(Asc(Mid(sText, i, 1))))
    Next
 
    fsDecode_Email = sTemp
End Function
Private Function fsReturnCharacter(ByVal iIndex As Integer) As String
    Dim vAlpha(71) As Variant
 
    vAlpha(0) = "A"
    vAlpha(1) = "B"
    vAlpha(2) = "C"
    vAlpha(3) = "D"
    vAlpha(4) = "E"
    vAlpha(5) = "F"
    vAlpha(6) = "G"
    vAlpha(7) = "H"
    vAlpha(8) = "I"
    vAlpha(9) = "J"
    vAlpha(10) = "K"
    vAlpha(11) = "L"
    vAlpha(12) = "M"
    vAlpha(13) = "N"
    vAlpha(14) = "O"
    vAlpha(15) = "P"
    vAlpha(16) = "Q"
    vAlpha(17) = "R"
    vAlpha(18) = "S"
    vAlpha(19) = "T"
    vAlpha(20) = "U"
    vAlpha(21) = "V"
    vAlpha(22) = "W"
    vAlpha(23) = "X"
    vAlpha(24) = "Y"
    vAlpha(25) = "Z"
    vAlpha(26) = " " 'space
    vAlpha(27) = "." 'dot
    vAlpha(28) = "-" ' hyhpen
    vAlpha(29) = "," ' coma
    vAlpha(30) = 0
    vAlpha(31) = 1
    vAlpha(32) = 2
    vAlpha(33) = 3
    vAlpha(34) = 4
    vAlpha(35) = 5
    vAlpha(36) = 6
    vAlpha(37) = 7
    vAlpha(38) = 8
    vAlpha(39) = 9
    vAlpha(40) = "/"
    vAlpha(41) = "a"
    vAlpha(42) = "b"
    vAlpha(43) = "c"
    vAlpha(44) = "d"
    vAlpha(45) = "e"
    vAlpha(46) = "f"
    vAlpha(47) = "g"
    vAlpha(48) = "h"
    vAlpha(49) = "i"
    vAlpha(50) = "j"
    vAlpha(51) = "k"
    vAlpha(52) = "l"
    vAlpha(53) = "m"
    vAlpha(54) = "n"
    vAlpha(55) = "o"
    vAlpha(56) = "p"
    vAlpha(57) = "q"
    vAlpha(58) = "r"
    vAlpha(59) = "s"
    vAlpha(60) = "t"
    vAlpha(61) = "u"
    vAlpha(62) = "v"
    vAlpha(63) = "w"
    vAlpha(64) = "x"
    vAlpha(65) = "y"
    vAlpha(66) = "z"
    vAlpha(67) = ":"
    vAlpha(68) = "@"
    vAlpha(69) = "_"
    vAlpha(70) = "("
    vAlpha(71) = ")"
    
    fsReturnCharacter = vAlpha(iIndex)
End Function
Private Function fiReturnCityIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(26) As Variant
    vAscii(0) = 194
    vAscii(1) = 195
    vAscii(2) = 196
    vAscii(3) = 197
    vAscii(4) = 198
    vAscii(5) = 199
    vAscii(6) = 200
    vAscii(7) = 201
    vAscii(8) = 202
    vAscii(9) = 203
    vAscii(10) = 204
    vAscii(11) = 205
    vAscii(12) = 206
    vAscii(13) = 207
    vAscii(14) = 208
    vAscii(15) = 209
    vAscii(16) = 210
    vAscii(17) = 211
    vAscii(18) = 212
    vAscii(19) = 213
    vAscii(20) = 214
    vAscii(21) = 215
    vAscii(22) = 216
    vAscii(23) = 217
    vAscii(24) = 218
    vAscii(25) = 219
    vAscii(26) = 161
       
        
    For iIndex = 0 To 26
        If iAscii = vAscii(iIndex) Then
             fiReturnCityIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnCityIndex = 25
End Function

Private Function fiReturnStateIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(26) As Variant
    vAscii(0) = 198
    vAscii(1) = 199
    vAscii(2) = 200
    vAscii(3) = 201
    vAscii(4) = 202
    vAscii(5) = 203
    vAscii(6) = 204
    vAscii(7) = 205
    vAscii(8) = 206
    vAscii(9) = 207
    vAscii(10) = 208
    vAscii(11) = 209
    vAscii(12) = 210
    vAscii(13) = 211
    vAscii(14) = 212
    vAscii(15) = 213
    vAscii(16) = 214
    vAscii(17) = 215
    vAscii(18) = 216
    vAscii(19) = 217
    vAscii(20) = 218
    vAscii(21) = 219
    vAscii(22) = 220
    vAscii(23) = 221
    vAscii(24) = 222
    vAscii(25) = 223
    vAscii(26) = 165
       
        
    For iIndex = 0 To 26
        If iAscii = vAscii(iIndex) Then
             fiReturnStateIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnStateIndex = 25
End Function
Private Function fiReturnContactPersonIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(27) As Variant
    vAscii(0) = 182
    vAscii(1) = 183
    vAscii(2) = 184
    vAscii(3) = 185
    vAscii(4) = 186
    vAscii(5) = 187
    vAscii(6) = 188
    vAscii(7) = 189
    vAscii(8) = 190
    vAscii(9) = 191
    vAscii(10) = 192
    vAscii(11) = 193
    vAscii(12) = 194
    vAscii(13) = 195
    vAscii(14) = 196
    vAscii(15) = 197
    vAscii(16) = 198
    vAscii(17) = 199
    vAscii(18) = 200
    vAscii(19) = 201
    vAscii(20) = 202
    vAscii(21) = 203
    vAscii(22) = 204
    vAscii(23) = 205
    vAscii(24) = 206
    vAscii(25) = 207
    vAscii(26) = 149
    vAscii(27) = 163 'dot
       
        
    For iIndex = 0 To 27
        If iAscii = vAscii(iIndex) Then
             fiReturnContactPersonIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnContactPersonIndex = 25
End Function
Private Function fiReturnCompanyIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(39) As Variant
    vAscii(0) = 184
    vAscii(1) = 185
    vAscii(2) = 186
    vAscii(3) = 187
    vAscii(4) = 188
    vAscii(5) = 189
    vAscii(6) = 190
    vAscii(7) = 191
    vAscii(8) = 192
    vAscii(9) = 193
    vAscii(10) = 194
    vAscii(11) = 195
    vAscii(12) = 196
    vAscii(13) = 197
    vAscii(14) = 198
    vAscii(15) = 199
    vAscii(16) = 200
    vAscii(17) = 201
    vAscii(18) = 202
    vAscii(19) = 203
    vAscii(20) = 204
    vAscii(21) = 205
    vAscii(22) = 206
    vAscii(23) = 207
    vAscii(24) = 208
    vAscii(25) = 209
    vAscii(26) = 151 ' space
    vAscii(27) = 165 'dot
    vAscii(28) = 164 ' -
    vAscii(29) = 163 ',
    vAscii(30) = 167 '0
    vAscii(31) = 168 '1
    vAscii(32) = 169 '2
    vAscii(33) = 170 '3
    vAscii(34) = 171 '4
    vAscii(35) = 172 '5
    vAscii(36) = 173 '6
    vAscii(37) = 174 '7
    vAscii(38) = 175 '8
    vAscii(39) = 176 '9
       
        
    For iIndex = 0 To 39
        If iAscii = vAscii(iIndex) Then
             fiReturnCompanyIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnCompanyIndex = 25
End Function
Private Function fiReturnProfilesIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(27) As Variant
    vAscii(0) = 186
    vAscii(1) = 187
    vAscii(2) = 188
    vAscii(3) = 189
    vAscii(4) = 190
    vAscii(5) = 191
    vAscii(6) = 192
    vAscii(7) = 193
    vAscii(8) = 194
    vAscii(9) = 195
    vAscii(10) = 196
    vAscii(11) = 197
    vAscii(12) = 198
    vAscii(13) = 199
    vAscii(14) = 200
    vAscii(15) = 201
    vAscii(16) = 202
    vAscii(17) = 203
    vAscii(18) = 204
    vAscii(19) = 205
    vAscii(20) = 206
    vAscii(21) = 207
    vAscii(22) = 208
    vAscii(23) = 209
    vAscii(24) = 210
    vAscii(25) = 211
    vAscii(26) = 153 ' space
    vAscii(27) = 167 'dot
    vAscii(29) = 165 ', comma
       
        
    For iIndex = 0 To 29
        If iAscii = vAscii(iIndex) Then
             fiReturnProfilesIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnProfilesIndex = 25
End Function
Private Function fiReturnAddress1Index(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(39) As Variant
    vAscii(0) = 188
    vAscii(1) = 189
    vAscii(2) = 190
    vAscii(3) = 191
    vAscii(4) = 192
    vAscii(5) = 193
    vAscii(6) = 194
    vAscii(7) = 195
    vAscii(8) = 196
    vAscii(9) = 197
    vAscii(10) = 198
    vAscii(11) = 199
    vAscii(12) = 200
    vAscii(13) = 201
    vAscii(14) = 202
    vAscii(15) = 203
    vAscii(16) = 204
    vAscii(17) = 205
    vAscii(18) = 206
    vAscii(19) = 207
    vAscii(20) = 208
    vAscii(21) = 209
    vAscii(22) = 210
    vAscii(23) = 211
    vAscii(24) = 212
    vAscii(25) = 213
    vAscii(26) = 155 ' space
    vAscii(27) = 169 'dot
    vAscii(28) = 168 ' -
    vAscii(29) = 167 ',
    vAscii(30) = 171 '0
    vAscii(31) = 172 '1
    vAscii(32) = 173 '2
    vAscii(33) = 174 '3
    vAscii(34) = 175 '4
    vAscii(35) = 176 '5
    vAscii(36) = 177 '6
    vAscii(37) = 178 '7
    vAscii(38) = 179 '8
    vAscii(39) = 180 '9
    
    
        
    For iIndex = 0 To 39
        If iAscii = vAscii(iIndex) Then
             fiReturnAddress1Index = iIndex
            Exit Function
        End If
    Next
    fiReturnAddress1Index = 25
End Function
Private Function fiReturnAddress2Index(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(39) As Variant
    vAscii(0) = 190
    vAscii(1) = 191
    vAscii(2) = 192
    vAscii(3) = 193
    vAscii(4) = 194
    vAscii(5) = 195
    vAscii(6) = 196
    vAscii(7) = 197
    vAscii(8) = 198
    vAscii(9) = 199
    vAscii(10) = 200
    vAscii(11) = 201
    vAscii(12) = 202
    vAscii(13) = 203
    vAscii(14) = 204
    vAscii(15) = 205
    vAscii(16) = 206
    vAscii(17) = 207
    vAscii(18) = 208
    vAscii(19) = 209
    vAscii(20) = 210
    vAscii(21) = 211
    vAscii(22) = 212
    vAscii(23) = 213
    vAscii(24) = 214
    vAscii(25) = 215
    vAscii(26) = 157 ' space
    vAscii(27) = 171 'dot
    vAscii(28) = 170 ' -
    vAscii(29) = 169 ',
    vAscii(30) = 173 '0
    vAscii(31) = 174 '1
    vAscii(32) = 175 '2
    vAscii(33) = 176 '3
    vAscii(34) = 177 '4
    vAscii(35) = 178 '5
    vAscii(36) = 179 '6
    vAscii(37) = 180 '7
    vAscii(38) = 181 '8
    vAscii(39) = 182 '9
    
    
        
    For iIndex = 0 To 39
        If iAscii = vAscii(iIndex) Then
             fiReturnAddress2Index = iIndex
            Exit Function
        End If
    Next
    fiReturnAddress2Index = 25
End Function

Private Function fiReturnAddress3Index(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(39) As Variant
    vAscii(0) = 192
    vAscii(1) = 193
    vAscii(2) = 194
    vAscii(3) = 195
    vAscii(4) = 196
    vAscii(5) = 197
    vAscii(6) = 198
    vAscii(7) = 199
    vAscii(8) = 200
    vAscii(9) = 201
    vAscii(10) = 202
    vAscii(11) = 203
    vAscii(12) = 204
    vAscii(13) = 205
    vAscii(14) = 206
    vAscii(15) = 207
    vAscii(16) = 208
    vAscii(17) = 209
    vAscii(18) = 210
    vAscii(19) = 211
    vAscii(20) = 212
    vAscii(21) = 213
    vAscii(22) = 214
    vAscii(23) = 215
    vAscii(24) = 216
    vAscii(25) = 217
    vAscii(26) = 159 ' space
    vAscii(27) = 173 'dot
    vAscii(28) = 172 ' -
    vAscii(29) = 171 ',
    vAscii(30) = 175 '0
    vAscii(31) = 176 '1
    vAscii(32) = 177 '2
    vAscii(33) = 178 '3
    vAscii(34) = 179 '4
    vAscii(35) = 180 '5
    vAscii(36) = 181 '6
    vAscii(37) = 182 '7
    vAscii(38) = 183 '8
    vAscii(39) = 184 '9
    
    
        
    For iIndex = 0 To 39
        If iAscii = vAscii(iIndex) Then
             fiReturnAddress3Index = iIndex
            Exit Function
        End If
    Next
    fiReturnAddress3Index = 25
End Function
Private Function fiReturnPinCodeIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(39) As Variant
   
   
    vAscii(30) = 179 '0
    vAscii(31) = 180 '1
    vAscii(32) = 181 '2
    vAscii(33) = 182 '3
    vAscii(34) = 183 '4
    vAscii(35) = 184 '5
    vAscii(36) = 185 '6
    vAscii(37) = 186 '7
    vAscii(38) = 187 '8
    vAscii(39) = 188 '9
    
    
        
    For iIndex = 30 To 39
        If iAscii = vAscii(iIndex) Then
             fiReturnPinCodeIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnPinCodeIndex = 25
End Function
Private Function fiReturnPhoneIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(40) As Variant
   
   
    vAscii(26) = 159 ' space
    vAscii(27) = 173 'dot
    vAscii(28) = 180 ' - or 180
    vAscii(29) = 171 ',
    vAscii(30) = 183 '0
    vAscii(31) = 184 '1
    vAscii(32) = 185 '2
    vAscii(33) = 186 '3
    vAscii(34) = 187 '4
    vAscii(35) = 188 '5
    vAscii(36) = 189 '6
    vAscii(37) = 190 '7
    vAscii(38) = 191 '8
    vAscii(39) = 192 '9
    vAscii(40) = 182 ' /
    
    
        
    For iIndex = 26 To 40
        If iAscii = vAscii(iIndex) Then
             fiReturnPhoneIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnPhoneIndex = 25
End Function
Private Function fiReturnFaxIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(40) As Variant
   
   
    vAscii(26) = 159 ' space
    vAscii(27) = 173 'dot
    vAscii(28) = 182 ' - or 180
    vAscii(29) = 171 ',
    vAscii(30) = 185 '0
    vAscii(31) = 186 '1
    vAscii(32) = 187 '2
    vAscii(33) = 188 '3
    vAscii(34) = 189 '4
    vAscii(35) = 190 '5
    vAscii(36) = 191 '6
    vAscii(37) = 192 '7
    vAscii(38) = 193 '8
    vAscii(39) = 194 '9
    vAscii(40) = 184 ' /
    
    
        
    For iIndex = 26 To 40
        If iAscii = vAscii(iIndex) Then
             fiReturnFaxIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnFaxIndex = 25
End Function
Private Function fiReturnWebsiteIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(67) As Variant
   
   
    vAscii(26) = 159 ' space
    vAscii(27) = 185 'dot
    vAscii(28) = 182 ' - or 180
    vAscii(29) = 171 ',
    vAscii(30) = 187 '0
    vAscii(31) = 188 '1
    vAscii(32) = 189 '2
    vAscii(33) = 190 '3
    vAscii(34) = 191 '4
    vAscii(35) = 192 '5
    vAscii(36) = 193 '6
    vAscii(37) = 194 '7
    vAscii(38) = 195 '8
    vAscii(39) = 196 '9
    vAscii(40) = 186 ' /
    vAscii(41) = 236
    vAscii(42) = 237
    vAscii(43) = 238
    vAscii(44) = 239
    vAscii(45) = 240
    vAscii(46) = 241
    vAscii(47) = 242
    vAscii(48) = 243
    vAscii(49) = 244
    vAscii(50) = 245
    vAscii(51) = 246
    vAscii(52) = 247
    vAscii(53) = 248
    vAscii(54) = 249
    vAscii(55) = 250
    vAscii(56) = 251
    vAscii(57) = 252
    vAscii(58) = 253
    vAscii(59) = 254
    vAscii(60) = 255
    vAscii(61) = 256
    vAscii(62) = 257
    vAscii(63) = 258
    vAscii(64) = 259
    vAscii(65) = 260
    vAscii(66) = 261
    vAscii(67) = 197 ' :
    
        
    For iIndex = 26 To 67
        If iAscii = vAscii(iIndex) Then
             fiReturnWebsiteIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnWebsiteIndex = 25
End Function
Private Function fiReturnEmailIndex(ByVal iAscii As Integer) As Integer
    Dim iIndex As Integer
    Dim vAscii(69) As Variant
      
    vAscii(26) = 159 ' space
    vAscii(27) = 161 'dot
    vAscii(28) = 182 ' - or 180
    vAscii(29) = 171 ',
    vAscii(30) = 163 '0
    vAscii(31) = 164 '1
    vAscii(32) = 165 '2
    vAscii(33) = 166 '3
    vAscii(34) = 167 '4
    vAscii(35) = 168 '5
    vAscii(36) = 169 '6
    vAscii(37) = 170 '7
    vAscii(38) = 171 '8
    vAscii(39) = 172 '9
    vAscii(40) = 186 ' /
    vAscii(41) = 212
    vAscii(42) = 213
    vAscii(43) = 214
    vAscii(44) = 215
    vAscii(45) = 216
    vAscii(46) = 217
    vAscii(47) = 218
    vAscii(48) = 219
    vAscii(49) = 220
    vAscii(50) = 221
    vAscii(51) = 222
    vAscii(52) = 223
    vAscii(53) = 224
    vAscii(54) = 225
    vAscii(55) = 226
    vAscii(56) = 227
    vAscii(57) = 228
    vAscii(58) = 229
    vAscii(59) = 230
    vAscii(60) = 231
    vAscii(61) = 232
    vAscii(62) = 233
    vAscii(63) = 234
    vAscii(64) = 235
    vAscii(65) = 236
    vAscii(66) = 237
    vAscii(67) = 197 ' :
    vAscii(68) = 179 '@
    vAscii(69) = 210 ' _

    For iIndex = 26 To 69
        If iAscii = vAscii(iIndex) Then
             fiReturnEmailIndex = iIndex
            Exit Function
        End If
    Next
    fiReturnEmailIndex = 25
End Function

