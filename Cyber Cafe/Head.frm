VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Head 
   BackColor       =   &H8000000C&
   Caption         =   "Cyber Master..."
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11160
   Icon            =   "Head.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "Head.frx":030A
   MousePointer    =   99  'Custom
   Picture         =   "Head.frx":045C
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   9000
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   6985
            Picture         =   "Head.frx":AB99
            Text            =   "Cyber Master (Version 1.0)"
            TextSave        =   "Cyber Master (Version 1.0)"
            Object.ToolTipText     =   "Alive Softwares Creation..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6985
            Text            =   "Alivesoftwares Creations..."
            TextSave        =   "Alivesoftwares Creations..."
            Key             =   "a"
            Object.ToolTipText     =   "Designed By vicky jadhav"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "11/3/2006"
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "11:36 AM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Align           =   1  'Align Top
      Height          =   1020
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   1799
      ButtonWidth     =   3810
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cyber &Master (1.0)"
            Key             =   "M"
            Object.ToolTipText     =   "Click here to Add customer inforamtion"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "c"
                  Text            =   "Add &Customer"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "p"
                  Text            =   "Get &Payment"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "x"
                  Text            =   "E&xit"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "All &Records"
            Key             =   "R"
            Object.ToolTipText     =   "Hit to see the daily records..."
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Show &Calinder"
            Key             =   "C"
            Object.ToolTipText     =   "Click to view Calinder"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&About"
            Key             =   "a"
            Object.ToolTipText     =   "To know about the software and also about programmer writer"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xit"
            Key             =   "x"
            Object.ToolTipText     =   "Close from here"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Computer Timer"
            Object.ToolTipText     =   "This Will show u since how long ur computer is on."
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Head.frx":AEB3
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   +{F1}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnucontent 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&Cyber Manager"
      End
   End
   Begin VB.Menu mnuwin 
      Caption         =   "&Arrange Window"
      Begin VB.Menu mnucas 
         Caption         =   "&Casced"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Head"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndplaysound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundName As String, ByVal uflags As Long) As Long




Private Sub MDIForm_Unload(Cancel As Integer)
sndplaysound (App.Path & "\1.wav"), 1
End Sub

Private Sub mnuabout_Click()
sndplaysound (App.Path & "\Click.wav"), 1
About.Show
Main.Hide
End Sub


Private Sub mnucas_Click()
    On Error GoTo ErrHandler
    Me.Arrange vbCascade
    Exit Sub
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Sub

End Sub

Private Sub mnuexit_Click()
sndplaysound (App.Path & "\1.wav"), 1
End
End Sub



Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
sndplaysound (App.Path & "\Click.wav"), 1
Main.Show


ElseIf Button.Index = 3 Then
sndplaysound (App.Path & "\Click.wav"), 1
Main.Hide
Record.Show



ElseIf Button.Index = 7 Then
sndplaysound (App.Path & "\Click.wav"), 1
Main.Hide
About.Show


ElseIf Button.Index = 5 Then
sndplaysound (App.Path & "\Click.wav"), 1
Main.Hide
Calinder.Show



ElseIf Button.Index = 9 Then
sndplaysound (App.Path & "\Click.wav"), 1
Unload Me

ElseIf Button.Index = 11 Then
sndplaysound (App.Path & "\Click.wav"), 1
Timer.Show
Main.Hide
End If
End Sub

Private Sub ToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

sndplaysound (App.Path & "\click.wav"), 1
If ButtonMenu.Index = 1 Then
Middle.Show

sndplaysound (App.Path & "\click.wav"), 1
ElseIf ButtonMenu.Index = 2 Then
Info.Show

sndplaysound (App.Path & "\click.wav"), 1
ElseIf ButtonMenu.Index = 3 Then
Unload Me

End If
End Sub
