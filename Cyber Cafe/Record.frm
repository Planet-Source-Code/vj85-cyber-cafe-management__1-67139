VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Record 
   Caption         =   "DAILY RECORD..."
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11085
   Icon            =   "Record.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Record.frx":030A
   ScaleHeight     =   9645
   ScaleWidth      =   11085
   WindowState     =   2  'Maximized
   Begin CYBERCAFE.xpButton XPCLO 
      Height          =   615
      Left            =   9480
      TabIndex        =   2
      Top             =   10440
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   1085
      TX              =   "&Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Onyx"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "Record.frx":9E98
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   10455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   18441
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColor       =   255
      ForeColor       =   16777215
      BackColorSel    =   0
      ForeColorSel    =   255
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      BorderStyle     =   0
      FormatString    =   $"Record.frx":A772
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CYBERCAFE.xpButton Command1 
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   10440
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   1058
      TX              =   "Show All  Records"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Onyx"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "Record.frx":A81C
   End
   Begin VB.Menu mnupntrecord 
      Caption         =   "&Print Record"
      Begin VB.Menu mnuprint 
         Caption         =   "&Print"
      End
   End
End
Attribute VB_Name = "Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Private Sub Command1_Click()
sndPlaySound (App.Path & "\Click.wav"), 1

Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset

MSFlexGrid1.Visible = False
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CAFE_DATABASE.mdb;Persist Security Info=False"
rs.Open "SELECT * FROM MASTER_TABLE", db, adOpenStatic, adLockReadOnly
rs.MoveFirst

MSFlexGrid1.Rows = rs.RecordCount + 1
MSFlexGrid1.Cols = rs.Fields.Count - 1
MSFlexGrid1.Row = 1
MSFlexGrid1.Col = 0
MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
MSFlexGrid1.Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
MSFlexGrid1.Row = 1
MSFlexGrid1.Visible = True

End Sub



Private Sub Form_Resize()
Dim Ctl As Control, CtlCln As New Collection
         On Error Resume Next
         For Each Ctl In Controls
            If Ctl.left < 0 Then CtlCln.Add Ctl
         Next
         ' Add the code to resize the controls:
         MSFlexGrid1.Move 0 * ScaleWidth, 0 * ScaleHeight, _
            1 * ScaleWidth, 0.9 * ScaleHeight
                            
        Command1.Move 0.001 * ScaleWidth, 0.9 * ScaleHeight, _
           0.6 * ScaleWidth, 0.1 * ScaleHeight
        
        XPCLO.Move 0.6 * ScaleWidth, 0.9 * ScaleHeight, _
           0.4 * ScaleWidth, 0.1 * ScaleHeight
        
      ' NOTE: The Height property can't be changed for the DriveListBox
      ' control or for the ComboBox control, whose Style property setting
      ' is 0 (Dropdown Combo) or 2 (Dropdown List). See the REFERENCES
      ' section for an article that discusses how to resize a ComboBox.
         For Each Ctl In CtlCln
            If Ctl.left > 0 Then Ctl.left = Ctl.left - 75000
         Next
         
End Sub

Private Sub Form_Unload(Cancel As Integer)
sndPlaySound (App.Path & "\1.wav"), 1
Main.Show
End Sub



Private Sub XPCLO_Click()
Unload Me
End Sub
