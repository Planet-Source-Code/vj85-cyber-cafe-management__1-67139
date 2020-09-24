VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Calinder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALINDER..."
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   Icon            =   "Calinder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Calinder.frx":030A
   ScaleHeight     =   7320
   ScaleWidth      =   7950
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   7320
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   12912
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Papyrus"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   8421504
      MultiSelect     =   -1  'True
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   20709377
      TitleBackColor  =   16777215
      TitleForeColor  =   255
      TrailingForeColor=   11183783
      CurrentDate     =   38928
   End
End
Attribute VB_Name = "Calinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndplaysound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundName As String, ByVal uflags As Long) As Long


Private Sub Form_Unload(Cancel As Integer)
sndplaysound (App.Path & "\1.wav"), 1
Main.Show
End Sub

Private Sub MonthView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
MonthView1.ToolTipText = Time & Date
End Sub
