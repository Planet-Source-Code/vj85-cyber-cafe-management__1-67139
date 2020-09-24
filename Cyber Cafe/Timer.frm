VERSION 5.00
Begin VB.Form Timer 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8880
   Icon            =   "Timer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Timer.frx":030A
   ScaleHeight     =   1185
   ScaleWidth      =   8880
   Begin VB.Timer Comtmr 
      Interval        =   1
      Left            =   0
      Top             =   2760
   End
   Begin VB.Label Label1 
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
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndplaysound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundName As String, ByVal uflags As Long) As Long
Private Declare Function GetTickCount& Lib "kernel32" ()
Private Sub ComTmr_Timer()
    Dim Secs, Mins, Hours, Days
    Dim TotalMins, TotalHours, TotalSecs, TempSecs
    Dim CaptionText
    TotalSecs = Int(GetTickCount / 1000)
    Days = Int(((TotalSecs / 60) / 60) / 24)
    TempSecs = Int(Days * 86400)
    TotalSecs = TotalSecs - TempSecs
    TotalHours = Int((TotalSecs / 60) / 60)
    TempSecs = Int(TotalHours * 3600)
    TotalSecs = TotalSecs - TempSecs
    TotalMins = Int(TotalSecs / 60)
    TempSecs = Int(TotalMins * 60)
    TotalSecs = (TotalSecs - TempSecs)
    If TotalHours > 23 Then
        Hours = (TotalHours - 23)
    Else
        Hours = TotalHours
    End If


    If TotalMins > 59 Then
        Mins = (TotalMins - (Hours * 60))
    Else
        Mins = TotalMins
    End If
    CaptionText = Days & " Days, " & Hours & " Hours, " & Mins & " Minutes, " & TotalSecs & " seconds" & vbCrLf
    Label1.Caption = CaptionText
    Me.Caption = CaptionText
End Sub


Private Sub Form_Unload(Cancel As Integer)
sndplaysound (App.Path & "\1.wav"), 1
Main.Show
End Sub
