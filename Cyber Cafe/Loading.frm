VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Loading 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOADING..."
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   Icon            =   "Loading.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Loading.frx":030A
   MousePointer    =   99  'Custom
   Picture         =   "Loading.frx":0614
   ScaleHeight     =   6390
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   0
      Top             =   4440
   End
   Begin MSComctlLib.ProgressBar PROGLOAD 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   6195
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "Loading.frx":45BE
      Scrolling       =   1
   End
End
Attribute VB_Name = "Loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long



Private Sub Form_Unload(Cancel As Integer)
    sndPlaySound (App.Path & "\1.wav"), 1

End Sub

' this will make ur progress bar Run
Private Sub Timer1_Timer()
    On Error GoTo Rani:
        With PROGLOAD
            .Value = .Value + 1
    End With
Exit Sub
Rani:
    If Err.Number = 380 Then
    sndPlaySound (App.Path & "\click.wav"), 1

        Unload Me
            Head.Show 'vbModal
            Main.Show
    End If
End Sub

