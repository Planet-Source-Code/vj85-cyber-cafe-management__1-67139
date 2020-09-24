VERSION 5.00
Begin VB.Form Middle 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMERS RECORDS & INFORMATIONS"
   ClientHeight    =   9465
   ClientLeft      =   3735
   ClientTop       =   1590
   ClientWidth     =   8055
   Icon            =   "Middle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Middle.frx":030A
   ScaleHeight     =   9465
   ScaleWidth      =   8055
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   7200
      Top             =   7080
   End
   Begin CYBERCAFE.xpButton CmdSave 
      Height          =   855
      Left            =   480
      TabIndex        =   28
      Top             =   8160
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1508
      TX              =   "SAVE CUSTOMER INFROMATION"
      ENAB            =   0   'False
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
      MICON           =   "Middle.frx":9E98
   End
   Begin CYBERCAFE.xpButton Cmdtime 
      Height          =   855
      Left            =   2040
      TabIndex        =   27
      Top             =   7200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1508
      TX              =   "CURRENT &TIME"
      ENAB            =   0   'False
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
      MICON           =   "Middle.frx":A772
   End
   Begin VB.CommandButton Sys1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   9
      Left            =   5520
      Picture         =   "Middle.frx":B04C
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Sys1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   8
      Left            =   4440
      Picture         =   "Middle.frx":B9E2
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Sys1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   7
      Left            =   3360
      Picture         =   "Middle.frx":C045
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Sys1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   6
      Left            =   2280
      Picture         =   "Middle.frx":C730
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Sys1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   5
      Left            =   1200
      Picture         =   "Middle.frx":CD7B
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Sys1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   4
      Left            =   5520
      Picture         =   "Middle.frx":D446
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Sys1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   3
      Left            =   4440
      Picture         =   "Middle.frx":DAC0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Sys1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   2
      Left            =   3360
      Picture         =   "Middle.frx":E155
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Sys1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   1
      Left            =   2280
      Picture         =   "Middle.frx":E80A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Sys1 
      BackColor       =   &H00000000&
      Height          =   1575
      Index           =   0
      Left            =   1200
      Picture         =   "Middle.frx":EF03
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtsysno 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   8160
      Width           =   150
   End
   Begin VB.TextBox txtnow 
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   8160
      Width           =   150
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1530
      Index           =   9
      Left            =   3360
      Picture         =   "Middle.frx":F4C6
      ScaleHeight     =   1530
      ScaleWidth      =   990
      TabIndex        =   13
      ToolTipText     =   "System numbers"
      Top             =   5040
      Width           =   990
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   8
      Left            =   3600
      Picture         =   "Middle.frx":FE5C
      ScaleHeight     =   1470
      ScaleWidth      =   480
      TabIndex        =   12
      ToolTipText     =   "System numbers"
      Top             =   5040
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1485
      Index           =   6
      Left            =   3600
      Picture         =   "Middle.frx":104BF
      ScaleHeight     =   1485
      ScaleWidth      =   615
      TabIndex        =   10
      ToolTipText     =   "System numbers"
      Top             =   5040
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   5
      Left            =   3600
      Picture         =   "Middle.frx":10B0A
      ScaleHeight     =   1455
      ScaleWidth      =   615
      TabIndex        =   9
      ToolTipText     =   "System numbers"
      Top             =   5040
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   4
      Left            =   3600
      Picture         =   "Middle.frx":111D5
      ScaleHeight     =   1455
      ScaleWidth      =   660
      TabIndex        =   8
      ToolTipText     =   "System numbers"
      Top             =   5040
      Width           =   660
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1425
      Index           =   3
      Left            =   3600
      Picture         =   "Middle.frx":1184F
      ScaleHeight     =   1425
      ScaleWidth      =   660
      TabIndex        =   7
      ToolTipText     =   "System numbers"
      Top             =   5040
      Width           =   660
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1425
      Index           =   2
      Left            =   3600
      Picture         =   "Middle.frx":11EE4
      ScaleHeight     =   1425
      ScaleWidth      =   600
      TabIndex        =   6
      ToolTipText     =   "System numbers"
      Top             =   5040
      Width           =   600
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1485
      Index           =   1
      Left            =   3600
      Picture         =   "Middle.frx":12599
      ScaleHeight     =   1485
      ScaleWidth      =   645
      TabIndex        =   5
      ToolTipText     =   "System numbers"
      Top             =   5040
      Width           =   645
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1485
      Index           =   0
      Left            =   3600
      Picture         =   "Middle.frx":12C92
      ScaleHeight     =   1485
      ScaleWidth      =   540
      TabIndex        =   4
      ToolTipText     =   "System numbers"
      Top             =   5040
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1425
      Index           =   7
      Left            =   3480
      Picture         =   "Middle.frx":13255
      ScaleHeight     =   1425
      ScaleWidth      =   765
      TabIndex        =   11
      ToolTipText     =   "System numbers"
      Top             =   5040
      Width           =   765
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   375
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Label lbltime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   6720
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   3375
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   1695
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblSys 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "System number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label lblSys 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the System number to allocate it"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label lblname 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   9255
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "Middle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This will play Wave files
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'This Connects ur Detabase
Dim dn As New ADODB.Connection
Dim rs As New ADODB.Recordset

'will scorl the form title
Dim x As Integer
Dim y As Integer
Dim prev As String


Private Sub CmdSave_Click()
    sndPlaySound (App.Path & "\click.wav"), 1

        If Len(txtName.text) = 0 Then
            MsgBox "Please enter the customer name", vbInformation, "Customer name !!!"
                
                ElseIf Len(txtsysno.text) = 0 Then
                MsgBox "Please Assign the system number", vbInformation, "System Number !!!"

            ElseIf Len(txtnow.text) = 0 Then
        MsgBox "Please assign the Incoming time", vbInformation, "In_Time !!!"

Else
    rs.AddNew
        rs.Fields(0).Value = txtName.text
            rs.Fields(1).Value = txtsysno.text
        rs.Fields(2).Value = txtnow.text
    rs.Update
    
        Unload Middle
    End If
End Sub

Private Sub CmdSave_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
sndPlaySound (App.Path & "\click.wav"), 1
Unload Me
End If

End Sub

Private Sub Cmdtime_Click()
    sndPlaySound (App.Path & "\click.wav"), 1

        lbltime.Caption = Time & Date
            txtnow.text = Time & Date
            CmdSave.Enabled = True
End Sub


Private Sub Cmdtime_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
sndPlaySound (App.Path & "\click.wav"), 1
Unload Me
End If

End Sub

Private Sub Form_Load()
dn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CAFE_DATABASE.mdb;Persist Security Info=False"
    rs.Open "select * from CURRENT_CUSTOMER", dn, adOpenDynamic, adLockOptimistic
    
        While rs.EOF <> True
             Sys1(rs.Fields(1).Value - 1).BackColor = &H808080
                Sys1(rs.Fields(1).Value - 1).Enabled = False
        rs.MoveNext
Wend
    
    Picture1(0).Visible = False
        Picture1(1).Visible = False
            Picture1(2).Visible = False
                Picture1(3).Visible = False
                    Picture1(4).Visible = False
                Picture1(5).Visible = False
            Picture1(6).Visible = False
        Picture1(7).Visible = False
    Picture1(8).Visible = False
Picture1(9).Visible = False

'this will scrole the form title
    prev = " CUSTOMER RECORD & INFORMATIONS... "
        x = Len(prev)
            y = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
sndPlaySound (App.Path & "\1.wav"), 1
dn.Close
Main.Show
End Sub


Private Sub Sys1_Click(Index As Integer)

sndPlaySound (App.Path & "\click.wav"), 1

txtsysno.text = Index + 1

If txtsysno.text = 1 Then
txtName.text = "Pc 1"
Picture1(0).Visible = True
    Picture1(1).Visible = False
        Picture1(2).Visible = False
            Picture1(3).Visible = False
                Picture1(4).Visible = False
                    Picture1(5).Visible = False
            Picture1(6).Visible = False
        Picture1(7).Visible = False
    Picture1(8).Visible = False
Picture1(9).Visible = False
Cmdtime.Enabled = True

ElseIf txtsysno.text = 2 Then
txtName.text = "Pc 2"
    Picture1(1).Visible = True
        Picture1(0).Visible = False
            Picture1(2).Visible = False
                Picture1(3).Visible = False
                    Picture1(4).Visible = False
                Picture1(5).Visible = False
            Picture1(6).Visible = False
        Picture1(7).Visible = False
    Picture1(8).Visible = False
Picture1(9).Visible = False
Cmdtime.Enabled = True

ElseIf txtsysno.text = 3 Then
txtName.text = "Pc 3"
    Picture1(2).Visible = True
        Picture1(0).Visible = False
            Picture1(1).Visible = False
                Picture1(3).Visible = False
                    Picture1(4).Visible = False
                Picture1(5).Visible = False
            Picture1(6).Visible = False
        Picture1(7).Visible = False
    Picture1(8).Visible = False
Picture1(9).Visible = False
Cmdtime.Enabled = True

ElseIf txtsysno.text = 4 Then
txtName.text = "Pc 4"
    Picture1(3).Visible = True
        Picture1(0).Visible = False
            Picture1(1).Visible = False
                Picture1(2).Visible = False
                    Picture1(4).Visible = False
                Picture1(5).Visible = False
            Picture1(6).Visible = False
        Picture1(7).Visible = False
    Picture1(8).Visible = False
Picture1(9).Visible = False
Cmdtime.Enabled = True

ElseIf txtsysno.text = 5 Then
txtName.text = "Pc 5"
    Picture1(4).Visible = True
        Picture1(0).Visible = False
            Picture1(1).Visible = False
                Picture1(2).Visible = False
                    Picture1(3).Visible = False
                Picture1(5).Visible = False
            Picture1(6).Visible = False
        Picture1(7).Visible = False
    Picture1(8).Visible = False
Picture1(9).Visible = False
Cmdtime.Enabled = True

ElseIf txtsysno.text = 6 Then
txtName.text = "Pc 6"
    Picture1(5).Visible = True
        Picture1(0).Visible = False
            Picture1(1).Visible = False
                Picture1(2).Visible = False
                    Picture1(3).Visible = False
                Picture1(4).Visible = False
            Picture1(6).Visible = False
        Picture1(7).Visible = False
    Picture1(8).Visible = False
Picture1(9).Visible = False
Cmdtime.Enabled = True

ElseIf txtsysno.text = 7 Then
txtName.text = "Pc 7"
    Picture1(6).Visible = True
        Picture1(0).Visible = False
            Picture1(1).Visible = False
                Picture1(2).Visible = False
                    Picture1(3).Visible = False
                Picture1(4).Visible = False
            Picture1(5).Visible = False
        Picture1(7).Visible = False
    Picture1(8).Visible = False
Picture1(9).Visible = False
Cmdtime.Enabled = True

ElseIf txtsysno.text = 8 Then
txtName.text = "Pc 8"
    Picture1(7).Visible = True
        Picture1(0).Visible = False
            Picture1(1).Visible = False
                Picture1(2).Visible = False
                    Picture1(3).Visible = False
                Picture1(4).Visible = False
            Picture1(5).Visible = False
        Picture1(6).Visible = False
    Picture1(8).Visible = False
Picture1(9).Visible = False
Cmdtime.Enabled = True

ElseIf txtsysno.text = 9 Then
txtName.text = "Pc 9"
    Picture1(8).Visible = True
        Picture1(0).Visible = False
            Picture1(1).Visible = False
                Picture1(2).Visible = False
                    Picture1(3).Visible = False
                Picture1(4).Visible = False
            Picture1(5).Visible = False
        Picture1(6).Visible = False
    Picture1(7).Visible = False
Picture1(9).Visible = False
Cmdtime.Enabled = True

ElseIf txtsysno.text = 10 Then
txtName.text = "Pc 10"
    Picture1(9).Visible = True
        Picture1(0).Visible = False
            Picture1(1).Visible = False
                Picture1(2).Visible = False
                    Picture1(3).Visible = False
                Picture1(4).Visible = False
            Picture1(5).Visible = False
        Picture1(6).Visible = False
    Picture1(7).Visible = False
Picture1(8).Visible = False
Cmdtime.Enabled = True

    End If
End Sub

Private Sub Sys1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
sndPlaySound (App.Path & "\click.wav"), 1
Unload Me
End If

End Sub

Private Sub Timer1_Timer()
    Middle.Caption = Mid$(prev, y, x)
        y = y + 1
            If y > x Then
        y = 1
    End If
End Sub



Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
sndPlaySound (App.Path & "\click.wav"), 1
Unload Me
End If
End Sub
