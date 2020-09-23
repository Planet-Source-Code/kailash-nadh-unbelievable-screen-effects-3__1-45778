VERSION 5.00
Begin VB.Form frmSet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unbelievable Effects v3"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2790
   Icon            =   "frmset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox cmb 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "About"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Play"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose an Effect:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1305
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code by Kailash Nadh , India
'kailash@bnsoft.net , http://bnsoft.net

'After about 8 months, the LATEST release. V 3


'Minimize all windows. Not actually necessary, but
'to see the real effect, the program should melt off the Desktop!
Private Declare Sub keybd_event Lib "USER32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B

Private Sub Command1_Click()
    Me.Tag = Me.cmb.ListIndex
    frmMain.doeffect
    'Set this form's tag to the effect number
End Sub

Private Sub Command3_Click()
    about.Show vbModal, Me
End Sub

Private Sub Form_Load()
    Me.Hide
    minAll ' Minimize all open windows
    
    'IMPORTANT!: DONOT!! KEEP THE SORTED
    'property of cmb to True!!
    
    
    cmb.AddItem "Melt" '0
    cmb.AddItem "Powder Blow" '1
    cmb.AddItem "Powder" '2
    cmb.AddItem "Evaporate" '3
    cmb.AddItem "Water Color" '4
    cmb.AddItem "Accumulate" '5
    cmb.AddItem "Checks" '6
    cmb.AddItem "Extreme Checks" '7
    cmb.AddItem "Wind Blow" '8
    cmb.AddItem "Pour Down" '9
    cmb.AddItem "Running Track" '10
    cmb.AddItem "Crazy Smoke" '11
    cmb.AddItem "Super Fast stream" '12
    cmb.AddItem "Moving water" '13
    cmb.AddItem "Powder+Water?" '14
    cmb.AddItem "Dissolve" '15
    cmb.AddItem "Blinds" '16
    cmb.AddItem "Stars" '17
    cmb.AddItem "Arrows" '18
    cmb.AddItem "Fire!" '19
    cmb.AddItem "Crushed+Grained" '20
    cmb.AddItem "Shake" '21

    cmb.ListIndex = 0

    Me.Show 'Finally, bring Screen Effects up!
    Me.WindowState = vbNormal
End Sub

Function minAll()
'###########################################
'Minimize All Windows to get the REAL EFFECT! ;)
Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(77, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Function
