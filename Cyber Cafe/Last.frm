VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Last 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALCULATE CUSTOMER PAYMENT"
   ClientHeight    =   6825
   ClientLeft      =   2190
   ClientTop       =   1785
   ClientWidth     =   9975
   Icon            =   "Last.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Last.frx":030A
   ScaleHeight     =   6825
   ScaleWidth      =   9975
   Begin VB.TextBox tXTDATE 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox TXTtIME 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   720
      Width           =   3015
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   9600
      Top             =   6360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6480
      Top             =   6840
   End
   Begin MSComctlLib.ProgressBar ProgCount 
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   4680
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin CYBERCAFE.xpButton XpCount 
      Height          =   735
      Left            =   240
      TabIndex        =   36
      Top             =   3840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1296
      TX              =   "Count Total Amount"
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
      MICON           =   "Last.frx":9E98
   End
   Begin VB.TextBox Txts 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   33
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox TxtP 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8760
      TabIndex        =   31
      Top             =   2520
      Width           =   975
   End
   Begin CYBERCAFE.xpButton XpScan 
      Height          =   735
      Left            =   8760
      TabIndex        =   27
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Onyx"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "Last.frx":A772
   End
   Begin CYBERCAFE.xpButton XpAmount 
      Height          =   735
      Left            =   8760
      TabIndex        =   26
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Onyx"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "Last.frx":B04C
   End
   Begin VB.TextBox txttotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Txt2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "5"
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Txt1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   8760
      TabIndex        =   22
      Text            =   "5"
      Top             =   2040
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6840
      Top             =   6840
   End
   Begin CYBERCAFE.xpButton Command2 
      Height          =   735
      Left            =   3960
      TabIndex        =   18
      Top             =   5760
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      TX              =   "&Close"
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
      MICON           =   "Last.frx":B926
   End
   Begin CYBERCAFE.xpButton Command1 
      Default         =   -1  'True
      Height          =   735
      Left            =   360
      TabIndex        =   17
      Top             =   5760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      TX              =   "Amount &Paid"
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
      MICON           =   "Last.frx":B942
   End
   Begin VB.ComboBox Text8 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Last.frx":B95E
      Left            =   3360
      List            =   "Last.frx":B971
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   525
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   510
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   3
      Left            =   3360
      TabIndex        =   43
      Text            =   "______________________________________"
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   0
      Left            =   3360
      TabIndex        =   40
      Text            =   " ---   "
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   42
      Text            =   "/"
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Txt 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   41
      Text            =   "-"
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   2
      Height          =   1095
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3255
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   2655
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   2655
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to add scanning Amt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   6720
      TabIndex        =   35
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to add Printing Amt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   0
      Left            =   6720
      TabIndex        =   34
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblprinting 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Scan's"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6720
      TabIndex        =   32
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblprinting 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Prints"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6720
      TabIndex        =   30
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblprinting 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning Amt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   29
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label lblprinting 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Amt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6720
      TabIndex        =   28
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rs"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   25
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblscan 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add Scanning's"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   21
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label lblprint 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add Printouts"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   20
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rs"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   19
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lblm 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5160
      TabIndex        =   16
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblh 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hh"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4080
      TabIndex        =   15
      Top             =   2880
      Width           =   270
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   975
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   6135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Received By"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Surfing Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Out_Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "In_Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System_Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   6615
      Left            =   120
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Last"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Dim dn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Public SYS_NO As String

'will scorl the form title
Dim x As Integer
Dim y As Integer
Dim prev As String


Private Sub Command1_Click()
sndPlaySound (App.Path & "\click.wav"), 1

If Len(Text8.text) = 0 Then
    MsgBox "Please Enter Receiver name", vbInformation, "Enter All the Details..."
Else
    rs.AddNew
    rs.Fields(0).Value = Text1.text
    rs.Fields(1).Value = Text2.text
    rs.Fields(2).Value = Text3.text
    rs.Fields(3).Value = Text4.text
    rs.Fields(4).Value = Text5.text
    rs.Fields(5).Value = Text7.text
    rs.Fields(6).Value = Text8.text
    rs.Fields(5).Value = txttotal.text
    rs.Fields(7).Value = tXTDATE.text
    rs.Fields(8).Value = Txt1.text
    rs.Fields(9).Value = TxtP.text
    rs.Fields(10).Value = XpAmount.Caption
    rs.Fields(11).Value = XpScan.Caption


    rs.Update
    RS1.Delete
    Main.Show
    Unload Me

End If



End Sub

Private Sub Command2_Click()
sndPlaySound (App.Path & "\1.wav"), 1
    
    Unload Me
    Info.Show 'vbModal
End Sub

Private Sub Form_Load()
dn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CAFE_DATABASE.mdb;Persist Security Info=False"
rs.Open "select * from MASTER_TABLE", dn, adOpenDynamic, adLockOptimistic
RS1.Open "select * from CURRENT_CUSTOMER WHERE SYSTEM_NO =" & SYS_NO, dn, adOpenDynamic, adLockOptimistic

'this will scrole the form title
prev = " TOTAL AMOUNT YOU HAVE TO PAY... "
x = Len(prev)
y = 1
TxtP.text = 0
Txts.text = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
sndPlaySound (App.Path & "\1.wav"), 1
dn.Close
End Sub

Private Sub Text6_LostFocus()
    If Val(Text9) <= 15 Then
        Text7 = 5
    ElseIf Val(Text9) <= 30 Then
        Text7 = 10
    ElseIf Val(Text9) <= 45 Then
        Text7 = 15
    ElseIf Val(Text9) <= 60 Then
        Text7 = 20
    Else
        Text7 = (Val(Text9) * 20) / 60
    End If
End Sub



Private Sub Timer1_Timer()
Last.Caption = Mid$(prev, y, x)
y = y + 1
If y > x Then
y = 1
End If
End Sub

Private Sub Timer2_Timer()
   On Error GoTo Rani:
        With ProgCount
            .Value = .Value + 1
    End With
Exit Sub
Rani:
    If Err.Number = 380 Then
    sndPlaySound (App.Path & "\click.wav"), 1
    Call Doing
    Text8.Enabled = True
    Command1.Enabled = True
    Timer2.Enabled = False
    End If
End Sub


Private Sub Timer3_Timer()
TXTtIME.text = Time
tXTDATE.text = Date
End Sub

Private Sub Txt1_Change()
Txt1.ToolTipText = "The Amount Has been Changed"
End Sub

Private Sub Txt1_Click()
Txt1.Locked = True
End Sub

Private Sub Txt1_DblClick()
Txt1.Locked = False
End Sub

Private Sub Txt2_Change()
Txt2.ToolTipText = "The Amount Has been Changed"
End Sub

Private Sub Txt2_Click()
Txt2.Locked = True
End Sub

Private Sub Txt2_DblClick()
Txt2.Locked = False
End Sub

'Printer Amt

Private Sub TxtP_Change()
On Error GoTo Maria
Dim Vicky As Integer
Dim Jenny As Integer
Vicky = Txt1.text
Jenny = TxtP.text
XpAmount.Caption = Vicky * Jenny
Maria:
If Err.Number = 13 Then
MsgBox "Please Enter the Number of prints u have taken", vbInformation, "Cyber Master (Version 1.0)"
End If
End Sub

'Scanner Amt

Private Sub Txts_Change()
On Error GoTo Vicky
Dim Britney As Integer
Dim Lopez As Integer
Britney = Txt2.text
Lopez = Txts.text
XpScan.Caption = Britney * Lopez
Vicky:
If Err.Number = 13 Then
MsgBox "Please Enter the Number of Scan's u have taken", vbInformation, "Cyber Master (Version 1.0)"
End If
End Sub

Private Sub xpButton2_Click()
On Error GoTo Vicky
Dim i As Integer
Dim L As Integer
i = Text7.text
L = Txt2.text
txttotal = i + L
Vicky:
If Err.Number = 13 Then
MsgBox "Please Enter Number or Prints and Scan", vbInformation, "Cyber Master"
End If
End Sub

Private Sub XpCount_Click()
Timer2.Enabled = True
ProgCount.Value = 0
End Sub

Private Sub Doing()
Dim Num1 As Integer
Dim Num2 As Integer
Dim Num3 As Integer
Num1 = Text7.text
Num2 = XpAmount.Caption
Num3 = XpScan.Caption
txttotal = Num1 + Num2 + Num3
End Sub
'Private Sub XpAmount_Click()
'Dim Kadie As Integer
'Dim Aviril As Integer
'Kadie = Text7.text
'Aviril = XpAmount.Caption
'txttotal = Kadie + Aviril
'End Sub

'Private Sub XpScan_Click()
'Dim Kadie As Integer
'Dim Aviril As Integer
'Kadie = Text7.text
'Aviril = XpAmount.Caption
'txttotal = Kadie + Aviril
'End Sub


