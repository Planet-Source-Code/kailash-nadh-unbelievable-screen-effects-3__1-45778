VERSION 5.00
Begin VB.Form about 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   DrawStyle       =   2  'Dot
   FillColor       =   &H0000FFFF&
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   2640
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "For more softwares && updates, visit my web site"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   570
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   2475
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.bnsoft.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kailash@bnsoft.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1605
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age:16 , India"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by Kailash Nadh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   2190
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unbelievable Screen Effects 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3405
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
rc = (Rnd * 100) / 100
gc = (Rnd * 100) / 150
bc = (Rnd * 100) / 100
Me.DrawWidth = Me.ScaleHeight / 255 * 2
For i = 0 To 255
Line (0, i * about.DrawWidth)-(about.DrawWidth * i, 0), RGB(rc * i, gc * i, bc * i)
Next i

For i = 0 To 255
Line (0, (i + 255) * Me.DrawWidth)-(Me.DrawWidth * (i + 255), 0), RGB(rc * (255 - i), gc * (255 - i), bc * (255 - i))
Next i

End Sub
