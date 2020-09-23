Attribute VB_Name = "modBevelMe"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Dim eCOL(10) As Single
Dim eCOL1(10) As Single
Dim eCOL2(10) As Single
Dim eCOL3(10) As Single
Dim eCOL4(10) As Single


Sub SizeOBJ(f, act)
Dim act2 As Integer
Select Case UCase(act)
Case "T": act2 = 12
Case "B": act2 = 15
Case "L": act2 = 10
Case "R": act2 = 11
Case "TL": act2 = 13
Case "TR": act2 = 14
Case "BL": act2 = 16
Case "BR": act2 = 17
Case Else
Exit Sub
End Select

Call ReleaseCapture
Call SendMessage(f.hwnd, &HA1, act2, 0)

End Sub
Sub DragOBJ(frm)

ReleaseCapture
SendMessage frm.hwnd, &HA1, 2, 0&

End Sub













Sub BevelMaster(pFRM, Col7, Maxi)

Dim gR As Single
Dim gG As Single
Dim gB As Single

Dim gR1 As Single
Dim gG1 As Single
Dim gB1 As Single

Dim gR2 As Single
Dim gG2 As Single
Dim gB2 As Single

Dim gR3 As Single
Dim gG3 As Single
Dim gB3 As Single

Dim gR4 As Single
Dim gG4 As Single
Dim gB4 As Single

Dim REScol1 As Long
Dim REScol2 As Long
Dim REScol3 As Long
Dim REScol4 As Long

Dim w1 As Integer
Dim h1 As Integer

Dim i7 As Integer
Dim gCOL As Long

gCOL = Col7

gR = ExtractR(gCOL)
gG = ExtractG(gCOL)
gB = ExtractB(gCOL)

gR1 = gR
gG1 = gG
gB1 = gB

eCOL(0) = 0.4
eCOL(1) = 0.7
eCOL(2) = 0.8
eCOL(3) = 0.9
eCOL(4) = 0.95

Dim diff As Single
diff = 0.05
eCOL1(0) = 0.8 - diff
eCOL1(1) = 1.12 - diff
eCOL1(2) = 1.11 - diff
eCOL1(3) = 1.07 - diff
eCOL1(4) = 1.08 - diff

eCOL2(0) = 0.8 - diff
eCOL2(1) = 1.2 - diff
eCOL2(2) = 1.15 - diff
eCOL2(3) = 1.1 - diff
eCOL2(4) = 1.08 - diff

eCOL3(0) = 0.5
eCOL3(1) = 0.72
eCOL3(2) = 0.84
eCOL3(3) = 0.9
eCOL3(4) = 0.95

eCOL4(0) = 0.4
eCOL4(1) = 0.7
eCOL4(2) = 0.8
eCOL4(3) = 0.9
eCOL4(4) = 0.95










w1 = pFRM.Width / 15 - 1
h1 = pFRM.Height / 15 - 1

pFRM.Cls
pFRM.BackColor = Col7

If Maxi = 0 Then
   pFRM.Line (0, 0)-(w1 * 15, h1 * 15), 0, B
   
Else

For i7 = 0 To Maxi - 1
   
   gR1 = gR * eCOL1(i7)
   gG1 = gG * eCOL1(i7)
   gB1 = gB * eCOL1(i7)
   REScol1 = RGB(gR1, gG1, gB1)
  
   gR2 = gR * eCOL2(i7)
   gG2 = gG * eCOL2(i7)
   gB2 = gB * eCOL2(i7)
   REScol2 = RGB(gR2, gG2, gB2)
  
   gR3 = gR * eCOL3(i7)
   gG3 = gG * eCOL3(i7)
   gB3 = gB * eCOL3(i7)
   REScol3 = RGB(gR3, gG3, gB3)
   
   gR4 = gR * eCOL(i7)
   gG4 = gG * eCOL(i7)
   gB4 = gB * eCOL(i7)
   REScol4 = RGB(gR4, gG4, gB4)
   
   pFRM.Line (i7 * 15, (h1 - i7) * 15)-(i7 * 15, i7 * 15), REScol1
   pFRM.Line (i7 * 15, i7 * 15)-((w1 - i7) * 15, i7 * 15), REScol2
   pFRM.Line ((w1 - i7) * 15, i7 * 15)-((w1 - i7) * 15, (h1 - i7) * 15), REScol3
   pFRM.Line ((w1 - i7) * 15, (h1 - i7) * 15)-(i7 * 15, (h1 - i7) * 15), REScol4

Next i7



End If


pFRM.Refresh

End Sub


Function ExtractR(ByVal CColor As Long) As Byte
ExtractR = CColor And 255
End Function
Function ExtractG(ByVal CColor As Long) As Byte
ExtractG = (CColor \ 256) And 255
End Function
Function ExtractB(ByVal CColor As Long) As Byte
ExtractB = (CColor \ 65536) And 255
End Function


