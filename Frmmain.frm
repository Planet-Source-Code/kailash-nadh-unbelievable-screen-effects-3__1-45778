VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code by Kailash Nadh , India
'kailash@bnsoft.net , http://bnsoft.net

'After about 8 months, the LATEST release. V 3


Option Explicit
Private Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "USER32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private lngDC As Long
Private blnLoop As Boolean
Dim m1 As Integer, m2 As Integer, z1 As Integer, z2 As Integer
Dim k1 As Integer, k2 As Integer

Private Sub Form_Click()
blnLoop = vbFalse
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
blnLoop = vbFalse
End Sub

Sub doeffect()
Dim intX As Integer, intY As Integer
Dim intI As Integer, intJ As Integer
Dim intWidth As Integer, intHeight As Integer
intWidth = Screen.Width / Screen.TwipsPerPixelX
intHeight = Screen.Height / Screen.TwipsPerPixelY
frmMain.Width = Screen.Width
frmMain.Height = Screen.Height
lngDC = GetDC(0)
Call BitBlt(hDC, 0, 0, intWidth, intHeight, lngDC, 0, 0, vbSrcCopy)
frmMain.Visible = vbTrue
frmMain.AutoRedraw = vbFalse
Randomize
blnLoop = vbTrue
Do While blnLoop = vbTrue
intX = (intWidth - 128) * Rnd
intY = (intHeight - 128) * Rnd
intI = m1 * Rnd - z1 ' The changes in values of m1 & m2 decides the effect
intJ = m2 * Rnd - z2
Call BitBlt(frmMain.hDC, intX + intI, intY + intJ, k1, k2, frmMain.hDC, intX, intY, vbSrcCopy)
DoEvents
Loop
Set frmMain = Nothing
Call ReleaseDC(0, lngDC)
Unload Me
End Sub

Private Sub Form_Load()
Dim ef As Integer
ef = frmSet.Tag 'Load the effect number from the frmSet tag
z1 = 1: z2 = 1: k1 = 128: k2 = 128
Select Case ef
Case 0
m1 = 2: m2 = 2

Case 1
m1 = 20: m2 = 20

Case 2
m1 = 9: m2 = 9

Case 3
m1 = 0: m2 = 0

Case 4
m1 = 3: m2 = 3

Case 5
m1 = 5: m2 = 5

Case 6
m1 = 10000: m2 = 10000

Case 7
m1 = 1000: m2 = 1000

Case 8
m1 = 10: m2 = 2

Case 9
m1 = 2: m2 = 10

Case 10
z1 = 10: z2 = 10

Case 11
z1 = 10: z2 = 10
m1 = 20: m2 = 10

Case 12
m1 = 2: m2 = 2
z1 = -100: z2 = 2

Case 13
m1 = 2: m2 = 2
k1 = 100: k2 = 10

Case 14
m1 = 10: m2 = 8
k1 = 100: k2 = 10

Case 15
m1 = 50: m2 = 10
k1 = 1: k2 = 25
z1 = 80: z2 = 10

Case 16
m1 = 2: m2 = 10
k1 = 12: k2 = 1
z1 = 5: z2 = 10

Case 17
m1 = 1: m2 = 1
k1 = 1: k2 = 1
z1 = -2: z2 = 10

Case 18
m1 = 5: m2 = 5
k1 = 8: k2 = 4
z1 = -2: z2 = 10

Case 19
m1 = 2: m2 = 10
k1 = 200: k2 = 4
z1 = -2: z2 = 10

Case 20
m1 = 30: m2 = 30
k1 = 10: k2 = 10
z1 = 10: z2 = 10

Case 21
m1 = 25: m2 = 25
k1 = 25: k2 = 255
z1 = 250: z2 = 25

End Select
End Sub
