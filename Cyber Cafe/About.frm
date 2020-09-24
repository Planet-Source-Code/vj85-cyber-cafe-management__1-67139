VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   5130
   ClientLeft      =   5685
   ClientTop       =   5430
   ClientWidth     =   5685
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "About.frx":030A
   MousePointer    =   99  'Custom
   Picture         =   "About.frx":0BD4
   ScaleHeight     =   5130
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   4440
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5280
      Top             =   4440
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4800
      Top             =   4440
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4320
      Top             =   4440
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   120
      MouseIcon       =   "About.frx":A762
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4560
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3975
      Left            =   120
      MouseIcon       =   "About.frx":B02C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Dim str As String
'Dim b As Integer
'Dim i As Integer
'Dim PI, Radius, Radians As Double


'Private Sub Form_Load()
'PI = 3.14159265358979
'Radius = 1680
'str = "Cyber Master... By Vicky! vic_Xcali_ky@yahoo.com"
'b = Len(str)
'i = 1
'End Sub

'Private Sub Timer1_Timer()
'Me.Caption = Left(str, i)
'i = i + 1
'If i = b + 1 Then
'Timer1.Interval = 3000
'i = 0
'End If
'End Sub
Option Explicit


Dim k As Integer
Dim text As String
Dim Text2 As String
Dim a As String



Dim str As String
Dim b As Integer
Dim i As Integer
Dim PI, Radius, Radians As Double

Private Sub Form_Click()
sndPlaySound (App.Path & "\1.wav"), 1
Main.Show
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
sndPlaySound (App.Path & "\1.wav"), 1
Main.Show

Unload Me
End If

End Sub

Private Sub Form_Load()
'PREVENTS FROM MORE THAN ONE INSTANCE RUNNING
If App.PrevInstance = True Then End
'>>>>>>>>>>>>>>>>>>>>>

k = 0
text = "Programmer Writer" & vbCrLf & " Vicky J" & vbCrLf & " from India " & vbCrLf & "Cyber Master(Version 1.0)!" & vbCrLf & "Perfect Cyber Cafe Manager Programme Specially designed for Cyber Cafe's" & vbCrLf & " Vote me on"
Text2 = "Vic_xcali_ky@yahoo.com"

'SetWindowPos Me.hwnd, -1, Me.Left / 15, _
'             Me.Top / 15, Me.Width / 15, _
'             Me.Height / 15, &H10 Or &H40




PI = 3.14159265358979
Radius = 1680
str = "Cyber Master (Version 1.0)"
b = Len(str)
i = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
sndPlaySound (App.Path & "\1.wav"), 1
Main.Show
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub


Private Sub Timer1_Timer()
Me.Caption = left(str, i)
i = i + 1
If i = b + 1 Then
Timer1.Interval = 3000
i = 0
End If
End Sub

Private Sub Timer2_Timer()
k = k + 1
a = Mid(text, k, 1)
Label1.Caption = Label1.Caption + a
If k >= Len(text) Then
    Timer2.Enabled = False
    Timer3.Enabled = True
    a = ""
    k = 0
End If

End Sub

Private Sub Timer3_Timer()
k = k + 1
a = Mid(Text2, k, 1)
Label2.Caption = Label2.Caption + a
If k > Len(Text2) Then
    Timer3.Enabled = False
    Timer4.Enabled = True
End If


End Sub



