VERSION 5.00
Begin VB.Form Info 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information & Customers Payment.."
   ClientHeight    =   7305
   ClientLeft      =   2190
   ClientTop       =   2925
   ClientWidth     =   11160
   Icon            =   "Count.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Count.frx":030A
   ScaleHeight     =   7305
   ScaleWidth      =   11160
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   10440
      Top             =   6600
   End
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   1215
      Index           =   0
      Left            =   4440
      TabIndex        =   20
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      TX              =   "Count"
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
      MICON           =   "Count.frx":9E98
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1215
      Index           =   0
      Left            =   240
      MousePointer    =   12  'No Drop
      Picture         =   "Count.frx":A772
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1215
      Index           =   1
      Left            =   5520
      MousePointer    =   12  'No Drop
      Picture         =   "Count.frx":ABE6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1095
      Index           =   9
      Left            =   5520
      MousePointer    =   12  'No Drop
      Picture         =   "Count.frx":B18E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1095
      Index           =   8
      Left            =   240
      MousePointer    =   12  'No Drop
      Picture         =   "Count.frx":B910
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1215
      Index           =   6
      Left            =   240
      MousePointer    =   12  'No Drop
      Picture         =   "Count.frx":BE82
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1215
      Index           =   3
      Left            =   5520
      MousePointer    =   12  'No Drop
      Picture         =   "Count.frx":C351
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1215
      Index           =   2
      Left            =   240
      MousePointer    =   12  'No Drop
      Picture         =   "Count.frx":C881
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1215
      Index           =   4
      Left            =   240
      MousePointer    =   12  'No Drop
      Picture         =   "Count.frx":CE51
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1215
      Index           =   5
      Left            =   5520
      MousePointer    =   12  'No Drop
      Picture         =   "Count.frx":D3D1
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1215
      Index           =   7
      Left            =   5520
      MousePointer    =   12  'No Drop
      Picture         =   "Count.frx":D92E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   1215
      Index           =   1
      Left            =   9840
      TabIndex        =   21
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      TX              =   "Count"
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
      MICON           =   "Count.frx":DEE9
   End
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   1215
      Index           =   2
      Left            =   4440
      TabIndex        =   22
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      TX              =   "Count"
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
      MICON           =   "Count.frx":E7C3
   End
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   1215
      Index           =   3
      Left            =   9840
      TabIndex        =   23
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      TX              =   "Count"
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
      MICON           =   "Count.frx":F09D
   End
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   1215
      Index           =   4
      Left            =   4440
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      TX              =   "Count"
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
      MICON           =   "Count.frx":F977
   End
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   1215
      Index           =   5
      Left            =   9840
      TabIndex        =   25
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      TX              =   "Count"
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
      MICON           =   "Count.frx":10251
   End
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   1215
      Index           =   6
      Left            =   4440
      TabIndex        =   26
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      TX              =   "Count"
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
      MICON           =   "Count.frx":10B2B
   End
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   1215
      Index           =   7
      Left            =   9840
      TabIndex        =   27
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      TX              =   "Count"
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
      MICON           =   "Count.frx":11405
   End
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   1095
      Index           =   8
      Left            =   4440
      TabIndex        =   28
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      TX              =   "Count"
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
      MICON           =   "Count.frx":11CDF
   End
   Begin CYBERCAFE.xpButton xpButton1 
      Height          =   1095
      Index           =   9
      Left            =   9840
      TabIndex        =   29
      Top             =   6000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      TX              =   "Count"
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
      MICON           =   "Count.frx":125B9
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   6720
      TabIndex        =   19
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   1320
      TabIndex        =   18
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   6720
      TabIndex        =   17
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   1320
      TabIndex        =   16
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   6720
      TabIndex        =   15
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   1320
      TabIndex        =   14
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   6720
      TabIndex        =   13
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   1320
      TabIndex        =   12
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   6720
      TabIndex        =   11
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   1320
      TabIndex        =   10
      Top             =   360
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Index           =   0
      Left            =   1200
      Top             =   240
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1095
      Index           =   9
      Left            =   6600
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1095
      Index           =   8
      Left            =   1200
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Index           =   7
      Left            =   6600
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Index           =   6
      Left            =   1200
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Index           =   5
      Left            =   6600
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Index           =   4
      Left            =   1200
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Index           =   3
      Left            =   6600
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Index           =   2
      Left            =   1200
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Index           =   1
      Left            =   6600
      Top             =   240
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      Height          =   7095
      Left            =   120
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Dim dn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS1_COUNT As New ADODB.Recordset
Dim in_time As Date
Dim out_time As Date
Dim h As Double
Dim m As Double
'SCORL FORM TITLE
Dim x As Integer
Dim y As Integer
Dim prev As String




Private Sub Command1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If

End Sub

'Private Sub Command2_Click(Index As Integer)
'RS1.Close
'RS1_COUNT.Close
'RS1.Open "select * from CURRENT_CUSTOMER WHERE SYSTEM_NO=" & Index + 1
'RS1_COUNT.Open "select count(*) from CURRENT_CUSTOMER WHERE SYSTEM_NO=" & Index + 1


'If RS1_COUNT.Fields(0).Value = 1 Then
'
'Last.SYS_NO = Index + 1



'in_time = RS1.Fields(2) .Value




'out_time = Now

'Last.Text1 = RS1.Fields(0).Value
'Last.Text2 = RS1.Fields(1).Value
'Last.Text3 = RS1.Fields(2).Value
'Last.Text4 = out_time

'Last.Text9 = Clear

'Last.Text9 = DateDiff("n", in_time, out_time)

  '  m = Val(Last.Text9) Mod 60
   ' h = (Last.Text9) / 60
    'h = Int(h)
    
    'Last.Text5.text = Val(h)
    'Last.Text9.text = Val(m)

    

''''''



'If DatePart("h", in_time) >= 22 Then
    'MsgBox DatePart("h", in_time)
    'If DatePart("h", in_time) = 24 Then
     '       If Val(Form3.Text9) <= 15 Then
      '          Form3.Text7 = 3.75
       '     ElseIf Val(Form3.Text9) <= 30 Then
       '         Form3.Text7 = 7.5
       '     ElseIf Val(Form3.Text9) <= 45 Then
       '         Form3.Text7 = 11.25
       '     ElseIf Val(Form3.Text9) <= 60 Then
       '         Form3.Text7 = 15
       '     Else
       '         Form3.Text7 = (Val(Form3.Text9) * 15) / 60
       '     End If
      'Form3.Text7 = Val(Form3.Text7) + (Val(Form3.Text5) * 15)
    'End If
    
'If DatePart("h", in_time) >= 0 Then
 '
  '  If DatePart("h", in_time) <= 6 Then
   '         If Val(Last.Text9) <= 15 Then
    '            Last.Text7 = 3.75
     '       ElseIf Val(Last.Text9) <= 30 Then
      '          Last.Text7 = 7.5
       '     ElseIf Val(Last.Text9) <= 45 Then
        '        Last.Text7 = 11.25
         '   ElseIf Val(Last.Text9) <= 60 Then
          ''      Last.Text7 = 15
            'Else
            '    Last.Text7 = (Val(Last.Text9) * 15) / 60
            'End If
            'Last.Text7 = Val(Last.Text7) + (Val(Last.Text5) * 15)
            
 '   ElseIf DatePart("h", in_time) <= 23 Then
  '          If Val(Last.Text9) <= 15 Then
   '             Last.Text7 = 5
    '        ElseIf Val(Last.Text9) <= 30 Then
     '           Last.Text7 = 10
      '      ElseIf Val(Last.Text9) <= 45 Then
       ''         Last.Text7 = 15
          '  ElseIf Val(Last.Text9) <= 60 Then
         '       Last.Text7 = 20
           ' Else
            '    Last.Text7 = (Val(Last.Text9) * 20) / 60
            'End If
            'Last.Text7 = Val(Last.Text7) + (Val(Last.Text5) * 20)
    'End If

'End If



 '   If Val(Form3.Text9) <= 15 Then
 '       Form3.Text7 = 5
 '   ElseIf Val(Form3.Text9) <= 30 Then
 '       Form3.Text7 = 10
 '   ElseIf Val(Form3.Text9) <= 45 Then
 '       Form3.Text7 = 15
 '   ElseIf Val(Form3.Text9) <= 60 Then
 '       Form3.Text7 = 20
 '   Else
 '       Form3.Text7 = (Val(Form3.Text9) * 20) / 60
 '   End If

 '    Form3.Text7 = Val(Form3.Text7) + (Val(Form3.Text5) * 20)
    

'Unload Me
'Last.Show 'vbModal
'Else
 '   MsgBox "There is no Customer on that System...", vbInformation, "No Customer Found On that System"
    
'End If
'End Sub

Private Sub Form_Load()
  dn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CAFE_DATABASE.mdb;Persist Security Info=False"
   rs.Open "select * from CURRENT_CUSTOMER", dn, adOpenDynamic, adLockOptimistic
   RS1.Open "select * from CURRENT_CUSTOMER", dn, adOpenDynamic, adLockOptimistic
   RS1_COUNT.Open "select count(*) from current_customer", dn, adOpenDynamic, adLockOptimistic
   While rs.EOF <> True
         Label1(rs.Fields(1).Value - 1).Caption = rs.Fields(0).Value
         rs.MoveNext
    Wend
   
   'SCORL FORM TITLE FROM HERE
   prev = "Information & Collect Customers Payment.. "
x = Len(prev)
y = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    sndPlaySound (App.Path & "\1.wav"), 1
    dn.Close
End Sub

Private Sub Timer1_Timer()
Info.Caption = Mid$(prev, y, x)
y = y + 1
If y > x Then
y = 1
End If
End Sub

Private Sub xpButton1_Click(Index As Integer)
sndPlaySound (App.Path & "\click.wav"), 1

RS1.Close
RS1_COUNT.Close
RS1.Open "select * from CURRENT_CUSTOMER WHERE SYSTEM_NO=" & Index + 1
RS1_COUNT.Open "select count(*) from CURRENT_CUSTOMER WHERE SYSTEM_NO=" & Index + 1


If RS1_COUNT.Fields(0).Value = 1 Then

Last.SYS_NO = Index + 1



in_time = RS1.Fields(2).Value




out_time = Now

Last.Text1 = RS1.Fields(0).Value
Last.Text2 = RS1.Fields(1).Value
Last.Text3 = RS1.Fields(2).Value
Last.Text4 = out_time

Last.Text9 = Clear

Last.Text9 = DateDiff("n", in_time, out_time)

    m = Val(Last.Text9) Mod 60
    h = (Last.Text9) / 60
    h = Int(h)
    
    Last.Text5.text = Val(h)
    Last.Text9.text = Val(m)

    

''''''



'If DatePart("h", in_time) >= 22 Then
    'MsgBox DatePart("h", in_time)
    'If DatePart("h", in_time) = 24 Then
     '       If Val(Form3.Text9) <= 15 Then
      '          Form3.Text7 = 3.75
       '     ElseIf Val(Form3.Text9) <= 30 Then
       '         Form3.Text7 = 7.5
       '     ElseIf Val(Form3.Text9) <= 45 Then
       '         Form3.Text7 = 11.25
       '     ElseIf Val(Form3.Text9) <= 60 Then
       '         Form3.Text7 = 15
       '     Else
       '         Form3.Text7 = (Val(Form3.Text9) * 15) / 60
       '     End If
      'Form3.Text7 = Val(Form3.Text7) + (Val(Form3.Text5) * 15)
    'End If
    
If DatePart("h", in_time) >= 0 Then
    
    If DatePart("h", in_time) <= 6 Then
            If Val(Last.Text9) <= 15 Then
                Last.Text7 = 5
            ElseIf Val(Last.Text9) <= 30 Then
                Last.Text7 = 10
            ElseIf Val(Last.Text9) <= 45 Then
                Last.Text7 = 15
            ElseIf Val(Last.Text9) <= 60 Then
                Last.Text7 = 15
            Else
                Last.Text7 = (Val(Last.Text9) * 15) / 60
            End If
            Last.Text7 = Val(Last.Text7) + (Val(Last.Text5) * 15)
            
    ElseIf DatePart("h", in_time) <= 23 Then
            If Val(Last.Text9) <= 15 Then
                Last.Text7 = 5
            ElseIf Val(Last.Text9) <= 30 Then
                Last.Text7 = 10
            ElseIf Val(Last.Text9) <= 45 Then
                Last.Text7 = 15
            ElseIf Val(Last.Text9) <= 60 Then
                Last.Text7 = 15
            Else
                Last.Text7 = (Val(Last.Text9) * 15) / 60
            End If
            Last.Text7 = Val(Last.Text7) + (Val(Last.Text5) * 15)
    End If

End If



 '   If Val(Form3.Text9) <= 15 Then
 '       Form3.Text7 = 5
 '   ElseIf Val(Form3.Text9) <= 30 Then
 '       Form3.Text7 = 10
 '   ElseIf Val(Form3.Text9) <= 45 Then
 '       Form3.Text7 = 15
 '   ElseIf Val(Form3.Text9) <= 60 Then
 '       Form3.Text7 = 20
 '   Else
 '       Form3.Text7 = (Val(Form3.Text9) * 20) / 60
 '   End If

 '    Form3.Text7 = Val(Form3.Text7) + (Val(Form3.Text5) * 20)
    

Unload Me
Last.Show 'vbModal
Else
    MsgBox "There is no Customer on that System...", vbInformation, "Cyber Master (Version 1.0)"
    
End If
End Sub

Private Sub xpButton1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If

End Sub
