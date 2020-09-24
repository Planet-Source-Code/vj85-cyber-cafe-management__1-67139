VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6255
   ClientLeft      =   4905
   ClientTop       =   2925
   ClientWidth     =   5730
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "Main.frx":030A
   MousePointer    =   99  'Custom
   Picture         =   "Main.frx":045C
   ScaleHeight     =   6255
   ScaleWidth      =   5730
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   840
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1482
      TX              =   "All &Records"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Onyx"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "Main.frx":9FEA
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   5280
      Top             =   0
   End
   Begin CYBERCAFE.xpButton xPABOUT 
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1508
      TX              =   "ABO&UT"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Onyx"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "Main.frx":A8C4
   End
   Begin CYBERCAFE.xpButton CmdEnd 
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1508
      TX              =   "&CLOSE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Onyx"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "Main.frx":B19E
   End
   Begin CYBERCAFE.xpButton CmdGet 
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1508
      TX              =   "GET &PAYMENT"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Onyx"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "Main.frx":BA78
   End
   Begin CYBERCAFE.xpButton CmdAdd 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1508
      TX              =   "&ADD CUSTOMER"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Onyx"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "Main.frx":C352
   End
   Begin VB.Label Time1a 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Bodoni MT Poster Compressed"
         Size            =   48
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   120
      Top             =   5040
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   3975
      Index           =   0
      Left            =   120
      Top             =   960
      Width           =   5535
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'scorols form title
Dim x As Integer
Dim y As Integer
Dim prev As String

Private Sub Form_Unload(Cancel As Integer)
sndPlaySound (App.Path & "\1.wav"), 1
End Sub

Private Sub Timer1_Timer()
Time1a.Caption = Time
Main.Caption = Mid$(prev, y, x)
y = y + 1
If y > x Then
y = 1
End If
End Sub

Private Sub Form_Load()
prev = " Cyber Master version (1.0) Perfect management software! Alivesoftwares Creation.  "
x = Len(prev)
y = 1
End Sub

Private Sub CmdAdd_Click()
sndPlaySound (App.Path & "\click.wav"), 1
Middle.Show 'vbModal
Me.Hide
End Sub

Private Sub CmdEnd_Click()
sndPlaySound (App.Path & "\1.wav"), 1
Unload Me
End Sub

Private Sub CmdGet_Click()
sndPlaySound (App.Path & "\click.wav"), 1
Info.Show 'vbModal
End Sub

Private Sub xPABOUT_Click()
sndPlaySound (App.Path & "\click.wav"), 1
About.Show 'vbModal
Main.Hide
End Sub

Private Sub xpButton1_Click()
Record.Show
End Sub
